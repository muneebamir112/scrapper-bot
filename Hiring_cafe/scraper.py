import os
import sys
if not os.environ.get("JOBBOT_LAUNCHER_AUTH"):
    import ctypes
    ctypes.windll.user32.MessageBoxW(0, "Access Denied: This module must be run from the Job Bot Launcher.", "Security Alert", 0x10)
    sys.exit(1)
import asyncio
import re
import json
from urllib.parse import urlparse, parse_qs, urlencode, quote
from patchright.async_api import async_playwright
import openpyxl
import sys
import os
import gspread
from google.oauth2.service_account import Credentials
from datetime import date

if getattr(sys, 'frozen', False):
    CURRENT_DIR = os.path.dirname(sys.executable)
    LAUNCHER_DIR = CURRENT_DIR
else:
    CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = os.path.dirname(os.path.dirname(CURRENT_DIR))
    LAUNCHER_DIR = os.path.join(BASE_DIR, "Job-bot-launcher")

# Persistent Chrome profile for the stealth browser (same patchright approach
# GlassD's scraper uses), kept separate from GlassD's own profile so each
# platform's login/session state stays independent.
PROFILE_DIR = os.path.join(CURRENT_DIR, "profiles", "hiringcafe_profile")

# Same Google Sheet + service account GlassD's scraper uses, so job links from
# every platform land in one place (Sheet1: No | Company Name | Job Title |
# Location | Job Age | Job Link | Date Added | Source).
import os
from dotenv import load_dotenv
load_dotenv(os.path.join(LAUNCHER_DIR, ".env"))
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1Kva2y5-54LXBWMTzNk_xp524ZE7N-CWiqL3VGATUZVM")
SERVICE_ACCOUNT_FILE = os.path.join(LAUNCHER_DIR, "service_account.json")


def get_google_sheet():
    """Connect to Sheet1 of the shared Google Sheet. Returns None (and prints why)
    if the service account file is missing or the connection fails, so the
    scraper can keep running and just skip the Sheets sync."""
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        print(f"  Google Sheets sync disabled: {SERVICE_ACCOUNT_FILE} not found")
        return None
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)
        try:
            worksheet = spreadsheet.worksheet("Jobs")
        except gspread.exceptions.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title="Jobs", rows=1000, cols=20)
        return worksheet
    except Exception as e:
        print(f"  Google Sheets connection failed: {e}")
        return None


def find_next_sheet_slot(worksheet):
    company_col = worksheet.col_values(2)
    
    if len(company_col) == 0 or (len(company_col) == 1 and not company_col[0].strip()):
        headers = ["No.", "Company Name", "Job Title", "Location", "Job Age", "Job Link", "Date Posted", "Platform"]
        worksheet.update(range_name="A1:H1", values=[headers], value_input_option="USER_ENTERED")
        
        # Format headers to match the user's yellow, bold, centered style
        try:
            worksheet.format("A1:H1", {
                "backgroundColor": {"red": 0.99, "green": 0.89, "blue": 0.58},
                "textFormat": {"bold": True, "fontSize": 10},
                "horizontalAlignment": "CENTER",
            })
        except Exception as e:
            print(f"  Could not format headers: {e}")
        return 2, 1
        
    next_row = len(company_col) + 1
    for i in range(1, len(company_col)):
        if not company_col[i].strip():
            next_row = i + 1
            break
            
    try:
        no_col = worksheet.col_values(1)
        if next_row - 1 < len(no_col) and no_col[next_row - 1].isdigit():
            next_no = int(no_col[next_row - 1]) + 1
        else:
            next_no = max([int(x) for x in no_col[1:] if x.isdigit()] + [0]) + 1
    except:
        next_no = 1
        
    return next_row, next_no


def load_seen_links(worksheet):
    """Application links already logged in the sheet (column F, 'Job Link'),
    used to skip a job only if this exact link was already added before."""
    try:
        link_col = worksheet.col_values(6)
    except Exception as e:
        print(f"  Could not read existing job links from sheet: {e}")
        return set()
    return {v.strip() for v in link_col[1:] if v.strip()}


def load_seen_jobs(worksheet):
    """(Company, Job Title) pairs already in the sheet (columns B and C),
    normalized for case/whitespace. The same real job posted on Glassdoor,
    Hiring Cafe, and Jobgether gets a different apply link on each platform,
    so load_seen_links() alone can't catch that - this catches it by what
    the job actually is instead of by its (platform-specific) link."""
    try:
        company_col = worksheet.col_values(2)
        title_col = worksheet.col_values(3)
    except Exception as e:
        print(f"  Could not read existing company/job-title pairs from sheet: {e}")
        return set()
    seen = set()
    for company, title in zip(company_col[1:], title_col[1:]):
        company = company.strip().lower()
        title = title.strip().lower()
        if company and title:
            seen.add((company, title))
    return seen


def is_within_24h(age_text):
    """Best-effort check of a 'posted X ago' style string. Hour/minute
    granularity is always within 24h; day/week/month granularity is not
    (except '0 days ago'). Unrecognized/empty text is treated as within 24h
    rather than silently dropping a job we can't confidently place."""
    if not age_text:
        return False
    text = age_text.lower().strip()
    if "just" in text or "today" in text or "now" in text:
        return True
    # hiring.cafe's own "Posted X ago" text uses abbreviated units (1w, 4d,
    # 2mo, 5h) rather than full words, which this pattern didn't recognize
    # before - every such age fell through to the "unrecognized" default of
    # True, silently letting week/month-old postings past the 24h filter.
    match = re.search(r'(\d+)\s*(mo|months?|w|weeks?|d|days?|h|hrs?|hours?|m|mins?|minutes?)\b', text)
    if not match:
        return True
    value = int(match.group(1))
    unit = match.group(2)
    if unit in ("h", "hr", "hrs", "hour", "hours"):
        return value <= 24
    if unit in ("m", "min", "mins", "minute", "minutes"):
        return True
    if unit in ("d", "day", "days"):
        return value <= 1
    return False  # week/month


CAPTCHA_INDICATORS = [
    "captcha", "recaptcha", "hcaptcha", "are you a robot", "are-you-a-robot",
    "cloudflare", "checking your browser", "verify you are human",
    "unusual traffic", "access denied", "attention required", "bot detection",
    "please verify", "security check", "just a moment",
]


def looks_like_captcha(url, page_text):
    """Best-effort check for whether a captured apply link actually landed on
    a CAPTCHA/bot-check page instead of a real job application form."""
    haystack = f"{url} {page_text}".lower()
    return any(indicator in haystack for indicator in CAPTCHA_INDICATORS)


async def _wait_for_real_navigation(page, timeout_ms=8000):
    """A newly opened popup sits on about:blank for a moment until its
    actual navigation lands - waiting a fixed delay and then reading
    page.url risks reading it while it's still blank (which was ending up
    as the captured application link). Wait for the URL to actually change
    instead."""
    try:
        await page.wait_for_url(lambda url: url != "about:blank", timeout=timeout_ms)
    except Exception:
        pass
    try:
        await page.wait_for_load_state("domcontentloaded", timeout=2000)
    except Exception:
        pass


def make_link_formula(url):
    """Wrap a URL in a HYPERLINK() formula so it renders as a clickable link
    in the sheet (gspread's default FORMATTED_VALUE reads still return the
    plain URL text, so nothing downstream that reads the sheet is affected)."""
    escaped = url.replace('"', '""')
    return f'=HYPERLINK("{escaped}", "{escaped}")'


def sync_one_job_to_sheet(sheet, job, seen_links, seen_jobs, next_row, next_no):
    """Write a single job to the next open sheet row right after it's
    extracted, instead of waiting to batch-sync the whole list at the end -
    that way jobs already found survive even if a later job or the browser
    itself crashes mid-run. Returns the (possibly advanced) next_row/next_no
    for the following call."""
    link = (job.get("apply_link") or job.get("job_url") or "").strip()
    if not link or link in seen_links:
        return next_row, next_no
    company = (job.get("company") or "").strip().lower()
    title = (job.get("title") or "").strip().lower()
    job_key = (company, title)
    if company and title and job_key in seen_jobs:
        return next_row, next_no
    try:
        sheet_row = [
            str(next_no), job.get("company", ""), job.get("title", ""),
            job.get("location", ""), job.get("posted", ""), make_link_formula(link),
            date.today().strftime("%m/%d/%Y"), "Hiring Cafe",
        ]
        sheet.update(range_name=f"A{next_row}:H{next_row}", values=[sheet_row], value_input_option="USER_ENTERED")
        seen_links.add(link)
        seen_jobs.add(job_key)
        return next_row + 1, next_no + 1
    except Exception as e:
        print(f"  Google Sheets sync failed for '{job.get('title', '')}': {e}")
        return next_row, next_no


async def main():
    keyword = sys.argv[1] if len(sys.argv) > 1 else "Full stack developer"
    location = sys.argv[2] if len(sys.argv) > 2 else "United States"

    jobs = []

    async with async_playwright() as p:
        os.makedirs(PROFILE_DIR, exist_ok=True)
        context = await p.chromium.launch_persistent_context(
            PROFILE_DIR,
            channel="chrome",
            headless=False,
            no_viewport=True,
        )
        page = context.pages[0] if context.pages else await context.new_page()

        # Instead of trying to click through the UI (which frequently changes
        # and has a lot of A/B testing overlays), we construct the search URL
        # directly. This is much faster and completely bypasses the FTUE banners,
        # location selectors, and search boxes.
        search_state = {
            "locations": [
                {
                    "formatted_address": location,
                    "types": ["country"],
                    "geometry": {
                        "location": {
                            "lat": 31.558,
                            "lon": 74.35071
                        }
                    },
                    "id": "user_country",
                    "address_components": [
                        {
                            "long_name": location,
                            "short_name": "US" if "united states" in location.lower() else location[:2].upper(),
                            "types": ["country"]
                        }
                    ],
                    "options": {
                        "flexible_regions": []
                    },
                    "workplace_types": ["Remote"]
                }
            ],
            "searchQuery": keyword
        }
        
        encoded_state = quote(json.dumps(search_state))
        search_url = f"https://hiringcafe.com/classic?searchState={encoded_state}"
        
        for attempt in range(4):
            try:
                print("  Clearing local storage to prevent default saved modes...")
                await page.goto("https://hiringcafe.com/", timeout=60000)
                await page.evaluate("window.localStorage.clear(); window.sessionStorage.clear();")
                await page.wait_for_timeout(1000)

                print(f"  Navigating to search URL: {search_url}")
                await page.goto(search_url, wait_until="domcontentloaded", timeout=60000)
                await page.wait_for_timeout(5000)
                
                # 1. Open the filter modal by clicking the chip if it exists
                try:
                    chip_selector = "button.flex.items-center.space-x-4"
                    await page.wait_for_selector(chip_selector, timeout=5000)
                    await page.click(chip_selector)
                    await page.wait_for_timeout(2000)
                    
                    # 2. Inside the modal, ensure Remote is checked and others unchecked
                    await page.evaluate('''() => {
                        const buttons = Array.from(document.querySelectorAll('button, div[role="button"], label'));
                        
                        buttons.forEach(btn => {
                            const txt = (btn.textContent || '').trim().toLowerCase();
                            const isChecked = btn.getAttribute('aria-checked') === 'true' || 
                                              btn.getAttribute('data-state') === 'checked' || 
                                              btn.classList.contains('bg-pink-600') ||
                                              (btn.querySelector('input') && btn.querySelector('input').checked);
                            
                            if ((txt === 'onsite' || txt === 'on-site' || txt === 'hybrid' || txt === 'field') && isChecked) {
                                btn.click();
                            } else if (txt === 'remote' && !isChecked) {
                                btn.click();
                            }
                        });
                    }''')
                    await page.wait_for_timeout(1000)
                    
                    # 3. If there is a location input, set it to United States
                    await page.evaluate('''() => {
                        const locInput = document.querySelector('input[placeholder*="location" i], input[placeholder*="country" i]');
                        if (locInput) {
                            locInput.value = 'United States';
                            locInput.dispatchEvent(new Event('input', { bubbles: true }));
                        }
                    }''')
                    await page.wait_for_timeout(1000)
                    
                    # 4. Click Apply or close modal
                    await page.evaluate('''() => {
                        const applyBtn = Array.from(document.querySelectorAll('button')).find(b => (b.textContent || '').trim().toLowerCase() === 'apply');
                        if (applyBtn) {
                            applyBtn.click();
                        } else {
                            // Press escape if no apply button
                            document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
                        }
                    }''')
                    await page.wait_for_timeout(3000)
                    
                except Exception as e:
                    print("  Could not interact with filter modal (maybe it didn't appear). Error:", e)

                # Fallback: if the chip is STILL showing the wrong text after closing the modal, delete it
                await page.evaluate('''() => {
                    const chips = Array.from(document.querySelectorAll('div.bg-gray-200.rounded-md'));
                    chips.forEach(chip => {
                        const txt = chip.textContent || '';
                        if (txt.includes('Hybrid') || txt.includes('Onsite') || txt.includes('Field') || (!txt.includes('United States') && !txt.includes('Remote'))) {
                            const trash = chip.querySelector('svg.text-red-500');
                            if (trash) {
                                const btn = trash.closest('button');
                                if (btn) btn.click();
                            }
                        }
                    });
                }''')
                await page.wait_for_timeout(2000)
                break
            except Exception as e:
                if attempt == 3:
                    print("Network error: Could not reach hiringcafe.com after multiple attempts.")
                    return
                print(f"Network hiccup reaching hiringcafe.com, retrying ({attempt + 1}/3)...")
                await asyncio.sleep(2)

        # Hiring Cafe uses URL pagination. We'll collect up to 50 links.
        all_job_links = []
        page_num = 1
        
        print("  Collecting job links across pages...")
        while True:
            try:
                # Wait for job links to load
                try:
                    await page.wait_for_selector('a[href^="/job/"]', timeout=10000)
                except:
                    print(f"  No more jobs found on page {page_num}.")
                    break

                current_links_data = await page.eval_on_selector_all(
                    'a[target="_blank"][rel="noopener noreferrer"][href^="/job/"]',
                    '''els => els.map(e => {
                        let card = e.closest("div");
                        for (let i=0; i<3; i++) { if (card && card.parentElement) card = card.parentElement; }
                        return {href: e.href, text: card ? card.innerText : ""};
                    })'''
                )
                
                unique_links = []
                for item in current_links_data:
                    href = item['href']
                    if href not in unique_links:
                        text = item['text'].lower()
                        if "today" in text or "just now" in text:
                            unique_links.append(href)
                            continue
                        
                        match = re.search(r'(?:posted)?\s*(\d+)\s*(mo|months?|w|weeks?|d|days?|h|hrs?|hours?|m|mins?|minutes?)(?:\s*ago)?\b', text)
                        if match:
                            value = int(match.group(1))
                            unit = match.group(2)
                            if unit in ("h", "hr", "hrs", "hour", "hours") and value <= 24:
                                unique_links.append(href)
                            elif unit in ("m", "min", "mins", "minute", "minutes"):
                                unique_links.append(href)
                            elif unit in ("d", "day", "days") and value <= 1:
                                unique_links.append(href)
                
                if len(current_links_data) == 0:
                    break
                    
                all_job_links.extend(unique_links)
                all_job_links = list(dict.fromkeys(all_job_links))
                print(f"  Page {page_num}: Found {len(unique_links)} fresh jobs out of {len(current_links_data)}. Total collected: {len(all_job_links)}")
                
                if len(all_job_links) >= 50:
                    print(f"  Collected enough links ({len(all_job_links)}). Stopping pagination.")
                    break

                # Prepare for next page
                current_url = page.url
                if 'page=' in current_url:
                    base_url = current_url.split('&page=')[0]
                else:
                    base_url = current_url
                next_url = f"{base_url}&page={page_num + 1}"
                
                await page.wait_for_timeout(2000)
                await page.goto(next_url, wait_until="domcontentloaded", timeout=60000)
                await page.wait_for_timeout(3000)
                page_num += 1
                
            except Exception as e:
                print(f"  Page {page_num} failed ({e}); stopping pagination...")
                break

        # Remove duplicates across pages
        all_job_links = list(dict.fromkeys(all_job_links))
        total_jobs = len(all_job_links)
        print(f"\nFound {total_jobs} unique job cards across all pages")

        if total_jobs == 0:
            print("No jobs found. Exiting.")
            return

        # Ask user how many to extract — only when actually run from an
        # interactive terminal. Without this check, input() blocks forever
        # when launched from the GUI (no console attached to read from),
        # which was hanging this scraper indefinitely.
        #
        # sys.stdin.isatty() isn't reliable when launched via
        # subprocess.Popen(stdin=DEVNULL) as the GUI does - it was sometimes
        # still reporting True there, so input() got attempted, immediately
        # failed (no stdin to read), and fell into the bare except below,
        # which used a leftover 5-job "just in case" cap instead of the
        # intended 25 - silently capping every GUI run at 5 jobs regardless
        # of how many were actually found. Track whether isatty() itself is
        # trustworthy separately from whether input() succeeded, so any
        # non-interactive/failure path lands on 25, not 5.
        try:
            interactive = sys.stdin.isatty()
        except Exception:
            interactive = False

        if interactive:
            try:
                num_to_extract = int(input(f"How many jobs to extract? (1-{total_jobs}): "))
                num_to_extract = min(num_to_extract, total_jobs, 25)  # Max 25 for safety
            except Exception:
                num_to_extract = min(25, total_jobs)
        else:
            num_to_extract = min(total_jobs, 25)

        # Connect once, before extraction starts, so each job below can be
        # written to the sheet as soon as it's scraped rather than waiting
        # to batch-sync the whole list at the end.
        sheet = get_google_sheet()
        sheet_next_row, sheet_next_no = find_next_sheet_slot(sheet) if sheet else (None, None)
        seen_links = load_seen_links(sheet) if sheet else set()
        seen_jobs = load_seen_jobs(sheet) if sheet else set()
        print(f"  Loaded {len(seen_links)} existing job links and {len(seen_jobs)} company/title pairs for duplicate check")

        print(f"\nExtracting {num_to_extract} jobs from the collected links...")
        for i, job_url in enumerate(all_job_links[:num_to_extract]):
            try:
                job_data = await extract_job_data(context, job_url)
            except Exception as e:
                print(f"    Error extracting {job_url}: {e}")
                job_data = None
            if job_data is not None:
                jobs.append(job_data)
                if sheet:
                    sheet_next_row, sheet_next_no = sync_one_job_to_sheet(
                        sheet, job_data, seen_links, seen_jobs, sheet_next_row, sheet_next_no
                    )

        await save_to_xlsx(jobs)
        try:
            await context.close()
        except Exception:
            pass
        print(f"\nDone! Saved {len(jobs)} jobs to jobs.xlsx")

async def extract_job_data(context, job_url):
    new_page = await context.new_page()
    try:
        await new_page.goto(job_url, wait_until="domcontentloaded", timeout=60000)
        await new_page.wait_for_timeout(5000)

        # hiring.cafe's page <title> is consistently formatted as
        # "<Job Title> at <Company> — <Location>", which is far less prone
        # to breaking on a redesign than matching specific CSS classes (the
        # previous selectors here no longer matched anything real on the
        # site's current layout - title was landing on the "Job description"
        # section heading, company and location were empty/wrong).
        title = ""
        company = ""
        job_location = ""
        page_title = await new_page.title()
        if " at " in page_title:
            title, rest = page_title.split(" at ", 1)
            title = title.strip()
            if " — " in rest:
                company, job_location = rest.rsplit(" — ", 1)
                company = company.strip()
                job_location = job_location.strip()
            else:
                company = rest.strip()
        if not title:
            h1_el = await new_page.query_selector('h1')
            if h1_el:
                title = (await h1_el.inner_text()).strip()

        posted = ""
        posted_el = await new_page.query_selector('div.text-xs.font-semibold.text-gray-500')
        if posted_el:
            posted = (await posted_el.inner_text()).replace("Posted", "").strip()

        if not is_within_24h(posted):
            await new_page.close()
            return None

        # A job is only worth saving if we can confirm a real, external,
        # ready-to-apply link - the hiring.cafe listing page itself isn't
        # one, so unlike before, this no longer falls back to job_url when
        # a genuine apply link can't be captured; it skips the job instead.
        def _is_ready_to_apply(link, text):
            if not link or link == "about:blank":
                return False
            if "hiring.cafe" in link:
                return False
            if looks_like_captcha(link, text):
                return False
            return True

        apply_link = None
        try:
            pages_before = len(context.pages)
            url_before_click = new_page.url

            # Click Apply button
            await new_page.evaluate('''() => {
                const buttons = document.querySelectorAll('button');
                for (let btn of buttons) {
                    if (btn.textContent && btn.textContent.includes('Apply')) {
                        btn.click();
                        return true;
                    }
                }
                return false;
            }''')

            # Wait for new page/tab to open
            await new_page.wait_for_timeout(6000)

            pages_after = len(context.pages)
            if pages_after > pages_before:
                # Apply opened a new tab - check that one.
                new_pages = context.pages[pages_before:]
                if new_pages:
                    await _wait_for_real_navigation(new_pages[0])
                    candidate_link = new_pages[0].url
                    try:
                        candidate_text = (await new_pages[0].title()) + " " + (await new_pages[0].inner_text("body"))[:500]
                    except Exception:
                        candidate_text = ""
                    await new_pages[0].close()
                    if _is_ready_to_apply(candidate_link, candidate_text):
                        apply_link = candidate_link
            else:
                # No new tab - some sites navigate the same page instead of
                # opening a popup. _wait_for_real_navigation() only waits out
                # an "about:blank" state, which doesn't apply here since this
                # page was already on a real hiring.cafe URL before the click
                # - compare against the pre-click URL instead to detect
                # whether it actually navigated anywhere.
                candidate_link = new_page.url
                if candidate_link != url_before_click:
                    try:
                        candidate_text = (await new_page.title()) + " " + (await new_page.inner_text("body"))[:500]
                    except Exception:
                        candidate_text = ""
                    if _is_ready_to_apply(candidate_link, candidate_text):
                        apply_link = candidate_link
        except Exception:
            pass

        await new_page.close()

        if not apply_link:
            return None

        return {"title": title.strip(), "company": company.strip(), "location": job_location.strip(),
                "posted": posted.strip(), "job_url": job_url, "apply_link": apply_link}
    except Exception as e:
        print(f"    Error: {e}")
        await new_page.close()
        return {"title": "", "company": "", "location": "", "posted": "", "job_url": job_url, "apply_link": ""}

async def save_to_xlsx(jobs):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Jobs"
    headers = ["title", "company", "location", "posted", "job_url", "apply_link"]
    ws.append(headers)
    for job in jobs:
        ws.append([job.get(h, "") for h in headers])
    wb.save("jobs.xlsx")

if __name__ == "__main__":
    asyncio.run(main())