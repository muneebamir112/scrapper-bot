from patchright.sync_api import sync_playwright, TimeoutError
from openpyxl import Workbook
import gspread
from google.oauth2.service_account import Credentials
import os
import sys

import os
import sys
if not os.environ.get("JOBBOT_LAUNCHER_AUTH"):
    msg = "Access Denied: This module must be run from the Job Bot Launcher."
    if sys.platform == "win32":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "Security Alert", 0x10)
    else:
        print(msg)
    sys.exit(1)
import sys
import re
from datetime import date
from dotenv import load_dotenv

if getattr(sys, 'frozen', False):
    CURRENT_DIR = os.path.dirname(sys.executable)
    LAUNCHER_DIR = CURRENT_DIR
else:
    CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = os.path.dirname(os.path.dirname(CURRENT_DIR))
    LAUNCHER_DIR = os.path.join(BASE_DIR, "Job-bot-launcher")

PROFILE_DIR = os.path.join(CURRENT_DIR, "profiles", "glassdoor_profile")

load_dotenv(os.path.join(LAUNCHER_DIR, ".env"))
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1Kva2y5-54LXBWMTzNk_xp524ZE7N-CWiqL3VGATUZVM")
SERVICE_ACCOUNT_FILE = os.path.join(LAUNCHER_DIR, "service_account.json")
EXCEL_HEADER = ["Company", "Job Title", "Location", "Job Age", "Application Link"]

def get_google_sheet():
    """Connect to the configured Google Sheet's first worksheet. Returns None (and
    prints why) if the service account file is missing or the connection fails, so
    the scraper can keep running and just skip the Sheets sync."""
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

def ensure_header(worksheet):
    """Create the header row if the tab is empty (column I holds the job description)."""
    company_col = worksheet.col_values(2)
    if len(company_col) == 0 or (len(company_col) == 1 and not company_col[0].strip()):
        headers = ["No.", "Company Name", "Job Title", "Location", "Job Age", "Job Link", "Date Posted", "Platform", "Job Description"]
        worksheet.update(range_name="A1:I1", values=[headers], value_input_option="USER_ENTERED")
        try:
            worksheet.format("A1:I1", {
                "backgroundColor": {"red": 0.99, "green": 0.89, "blue": 0.58},
                "textFormat": {"bold": True, "fontSize": 10},
                "horizontalAlignment": "CENTER",
            })
            worksheet.format("I2:I", {"numberFormat": {"type": "TEXT"}, "wrapStrategy": "CLIP"})
        except Exception as e:
            print(f"  Could not format headers: {e}")


def load_seen_links(worksheet):
    """Application links already logged in the Google Sheet (column F, 'Job Link').
    Used to skip a job only if this exact link was already added in a past run or
    earlier in the current one — multiple postings from the same company are fine
    as long as each has its own link."""
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

def parse_job_age(age_text):
    """Convert job age text to numeric days for comparison."""
    if not age_text:
        return 0
    age_text = age_text.lower().strip()
    if "today" in age_text or "just" in age_text or "now" in age_text:
        return 0
    # Match patterns: 23d, 5h, 2w, 1mo, 3mos, etc.
    match = re.search(r'(\d+)\s*(mo|mos|[dhmsw]|days?|hrs?|hours?|mins?|minutes?|weeks?|months?)', age_text)
    if not match:
        return 0 # Default to accepting if unrecognized, since fromAge=1 URL filter is applied
    
    value = int(match.group(1))
    unit = match.group(2)
    
    if unit in ('h', 'hr', 'hrs', 'hour', 'hours'):
        return value / 24
    elif unit in ('d', 'day', 'days'):
        return value
    elif unit in ('w', 'week', 'weeks'):
        return value * 7
    elif unit in ('m', 'mo', 'mos', 'month', 'months'):
        return value * 30
    elif unit in ('s', 'min', 'mins', 'minute', 'minutes'):
        return value / (24 * 60)
    return 0

CAPTCHA_INDICATORS = [
    "captcha", "recaptcha", "hcaptcha", "are you a robot", "are-you-a-robot",
    "cloudflare", "checking your browser", "verify you are human",
    "unusual traffic", "access denied", "attention required", "bot detection",
    "please verify", "security check", "just a moment",
]

def looks_like_captcha(url, page_text):
    """Best-effort check for whether a final application link actually landed
    on a CAPTCHA/bot-check page instead of a real job application form, so
    those links can be skipped rather than added to the sheet."""
    haystack = f"{url} {page_text}".lower()
    return any(indicator in haystack for indicator in CAPTCHA_INDICATORS)

def _click_and_follow(context, job_page, el):
    """Click el and return the page to read next: the new tab the click opened (the old page is
    closed), or job_page if the same tab navigated. Waits up to 15 s for the URL to leave glassdoor.com."""
    new_tab = None
    try:
        with context.expect_page(timeout=15000) as info:
            el.click()
            # Glassdoor may first ask the user to update job preferences; "Skip for now" dismisses
            # that prompt and the apply continues.
            try:
                skip = job_page.wait_for_selector("button:has-text('Skip for now')", timeout=4000)
                if skip and skip.is_visible():
                    skip.click()
            except TimeoutError:
                pass
        new_tab = info.value
    except TimeoutError:
        pass  # no new tab: the same page navigated (or nothing happened)
    if new_tab is not None:
        try:
            job_page.close()
        except Exception:
            pass
        job_page = new_tab
    try:
        job_page.wait_for_url(lambda u: u != "about:blank" and "glassdoor.com" not in u, timeout=15000)
    except TimeoutError:
        pass
    return job_page


def make_link_formula(url):
    """Wrap a URL in a HYPERLINK() formula so it renders as a clickable link
    in the sheet (gspread's default FORMATTED_VALUE reads still return the
    plain URL text, so nothing downstream that reads the sheet is affected)."""
    escaped = url.replace('"', '""')
    return f'=HYPERLINK("{escaped}", "{escaped}")'

import html as _html
import json
import time
from urllib.parse import urlparse


def is_within_24h(age_text):
    """Strict check of a 'posted X ago' text: only hours, minutes, 'today' and 'just now'
    count. Days, weeks, months and unrecognised text do not ('1d ago' can be ~47 h old)."""
    text = str(age_text or "").lower().strip()
    if not text:
        return False
    if "just now" in text or "today" in text or text == "now":
        return True
    match = re.search(r"(\d+)\+?\s*(h|hrs?|hours?|m|mins?|minutes?)\b", text)
    if not match:
        return False
    return match.group(2).startswith("m") or int(match.group(1)) <= 24


EMPTY_VALUES = {"", "n/a", "na", "unknown", "none", "null", "-", "undefined"}
BOARD_HOSTS = ("glassdoor.com", "hiring.cafe", "hiringcafe.com", "jobgether.com")

# Job-description capture: the page's JSON-LD first, then the largest block of readable text.
JSONLD_JS = """() => Array.from(document.querySelectorAll('script[type="application/ld+json"]')).map(s => s.textContent)"""
BLOCK_JS = """() => {
  const groups = ['[class*="description" i], [id*="description" i], [class*="job-details" i], [class*="jobdetail" i]', 'article', 'main, [role="main"]'];
  for (const g of groups) {
    let best = '';
    for (const el of document.querySelectorAll(g)) {
      const t = (el.innerText || '').trim();
      if (t.length > best.length && t.length < 60000) best = t;
    }
    if (best.length >= 400) return best;
  }
  return document.body ? document.body.innerText : '';
}"""


def html_to_text(raw):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", str(raw), flags=re.I | re.S)
    s = re.sub(r"<li\b[^>]*>", "\n- ", s, flags=re.I)
    s = re.sub(r"</?(p|div|br|ul|ol|h[1-6]|tr|section)\b[^>]*>", "\n", s, flags=re.I)
    return _html.unescape(re.sub(r"<[^>]+>", "", s))


def description_from_jsonld(raws):
    """Description of a schema.org JobPosting found in the page's JSON-LD blocks, as text."""
    def find(node):
        if isinstance(node, list):
            for item in node:
                found = find(item)
                if found:
                    return found
        elif isinstance(node, dict):
            kind = node.get("@type")
            if (kind == "JobPosting" or (isinstance(kind, list) and "JobPosting" in kind)) and isinstance(node.get("description"), str):
                return node["description"]
            for value in node.values():
                if isinstance(value, (list, dict)):
                    found = find(value)
                    if found:
                        return found
        return None

    for raw in raws or []:
        try:
            found = find(json.loads(raw))
        except Exception:
            continue
        if found:
            return html_to_text(found)
    return ""


def clean_description(text):
    """Make a description safe for one sheet cell: no control characters, tidy lines,
    capped at 20,000 characters, never starting with a character Sheets reads as a formula."""
    if not text:
        return ""
    s = str(text).replace("\r\n", "\n").replace("\r", "\n").replace("\xa0", " ")
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\u200b\u200c\u200d\u2060\ufeff]", "", s)
    s = "\n".join(re.sub(r"[ \t]+", " ", line).strip() for line in s.split("\n"))
    s = re.sub(r"\n{3,}", "\n\n", s).strip()
    if len(s) > 20000:
        s = s[:20000].rstrip() + "\n... [truncated]"
    if s[:1] in ("=", "+", "-", "@"):
        s = " " + s
    return s


def pick_description(candidates):
    """First candidate with at least 400 characters, else the longest non-empty one."""
    texts = [t for t in (clean_description(c) for c in candidates if c) if t]
    for t in texts:
        if len(t) >= 400:
            return t
    return max(texts, key=len) if texts else ""


def row_problem(job):
    """Why this job must not be written to the sheet, or '' if it is complete."""
    for field in ("company", "title", "location", "age", "link"):
        if str(job.get(field) or "").strip().lower() in EMPTY_VALUES:
            return f"missing {field}"
    link = str(job["link"]).strip()
    if not link.lower().startswith(("http://", "https://")):
        return "apply link is not an http(s) URL"
    host = (urlparse(link).netloc or "").lower()
    if any(host == h or host.endswith("." + h) for h in BOARD_HOSTS):
        return f"apply link is still on the job board ({host})"
    if looks_like_captcha(link, ""):
        return "apply link looks like a bot-check page"
    if not is_within_24h(job["age"]):
        return f"not posted within 24h ('{job['age']}')"
    return ""


def append_job(sheet, job, seen_links, seen_jobs):
    """Append one job to the sheet. Google picks the next free row, so scrapers running
    at the same time cannot overwrite each other. Skips duplicates and incomplete jobs.
    Returns True if the row was written."""
    link = str(job.get("link") or "").strip()
    key = (str(job.get("company") or "").strip().lower(), str(job.get("title") or "").strip().lower())
    if not link or link in seen_links or (key[0] and key[1] and key in seen_jobs):
        return False
    problem = row_problem(job)
    if problem:
        print(f"  Skipped '{job.get('title', '')}' @ {job.get('company', '')}: {problem}")
        return False
    row = [
        "=ROW()-1", job["company"], job["title"], job["location"], job["age"],
        make_link_formula(link), date.today().strftime("%m/%d/%Y"), job["platform"],
        clean_description(job.get("description", "")),
    ]
    for attempt in range(3):
        try:
            sheet.append_row(row, value_input_option="USER_ENTERED", table_range="A1")
            seen_links.add(link)
            seen_jobs.add(key)
            return True
        except Exception as e:
            if attempt == 2:
                print(f"  Google Sheets sync failed for '{job.get('title', '')}': {e}")
            else:
                time.sleep(2)
    return False


def read_description(page):
    """Description text of the page that is open: JSON-LD first, then the main text block."""
    candidates = []
    try:
        candidates.append(description_from_jsonld(page.evaluate(JSONLD_JS)))
    except Exception:
        pass
    try:
        candidates.append(page.evaluate(BLOCK_JS))
    except Exception:
        pass
    return pick_description(candidates)


def main():
    # For testing: limit number of jobs processed (set to None for unlimited)
    TEST_LIMIT = None  # set to integer for testing, None for all jobs  # reduced for quick test

    # Read inputs: prefer argv (job_title, location) since that's what the
    # launcher now passes for every run - inputs.txt is only a fallback for
    # running this script manually/standalone without arguments.
    if len(sys.argv) >= 3:
        job_title = sys.argv[1].strip()
        location = sys.argv[2].strip()
    else:
        try:
            with open("inputs.txt", "r") as f:
                lines = f.read().strip().splitlines()
                job_title = lines[0].strip()
                location = lines[1].strip()
        except Exception as e:
            print(f"Error reading inputs: {e}")
            return
    
    # Create Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Glassdoor Jobs"
    ws.append(EXCEL_HEADER)

    # Connect to Google Sheets (jobs are pushed there live, row by row, as they're scraped)
    sheet = get_google_sheet()
    if sheet:
        ensure_header(sheet)
    seen_links = load_seen_links(sheet) if sheet else set()
    seen_jobs = load_seen_jobs(sheet) if sheet else set()
    print(f"  Loaded {len(seen_links)} existing job links and {len(seen_jobs)} company/title pairs for duplicate check")

    # Launch stealth browser (real Chrome, persistent profile)
    os.makedirs(PROFILE_DIR, exist_ok=True)
    with sync_playwright() as p:
        try:
            context = p.chromium.launch_persistent_context(
                PROFILE_DIR,
                channel="chrome",
                headless=False,
                no_viewport=True,
            )
        except Exception as e:
            print(f"Launch failed: {e}")
            return

        page = context.pages[0] if context.pages else context.new_page()

        try:
            # Step 1: Navigate to Glassdoor Jobs
            for attempt in range(4):
                try:
                    page.goto("https://www.glassdoor.com/Job/index.htm", timeout=60000, wait_until="domcontentloaded")
                    break
                except Exception as e:
                    if attempt == 3:
                        print("Network error: Could not reach glassdoor.com after multiple attempts. Please check your internet connection.")
                        return
                    print(f"Network hiccup reaching glassdoor.com, retrying ({attempt + 1}/3)...")
                    page.wait_for_timeout(2000)
            page.wait_for_timeout(2000)
            
            # Close popups
            try:
                page.click("button[aria-label='Close'], button:has-text('Close')", timeout=3000)
            except TimeoutError:
                pass
            
            # Step 2: Fill search form
            try:
                page.wait_for_selector("input[aria-labelledby='searchBar-jobTitle_label'], input[name='sc.keyword']", timeout=15000)
                page.wait_for_selector("input[aria-labelledby='searchBar-location_label'], input[name='sc.location']", timeout=15000)
            except TimeoutError:
                print("  Search inputs not found, trying alternative selectors...")
                page.wait_for_selector("input[placeholder*='job title'], input[placeholder*='keyword']", timeout=15000)
                page.wait_for_selector("input[placeholder*='location'], input[name='location']", timeout=15000)
            
            # Fill job title
            page.fill("input[aria-labelledby='searchBar-jobTitle_label'], input[name='sc.keyword']", job_title)
            page.wait_for_timeout(300)
            page.press("input[aria-labelledby='searchBar-jobTitle_label'], input[name='sc.keyword']", "Escape")

            # Fill location
            page.fill("input[aria-labelledby='searchBar-location_label'], input[name='sc.location']", location)
            page.wait_for_timeout(3000)
            page.press("input[aria-labelledby='searchBar-location_label'], input[name='sc.location']", "Escape")

            # Submit search
            page.press("input[aria-labelledby='searchBar-location_label'], input[name='sc.location']", "Enter")
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except TimeoutError:
                pass

            # Close any popups
            try:
                page.click("button[aria-label='Close'], button:has-text('Close')", timeout=3000)
            except TimeoutError:
                pass

            # Step 4: Wait for results
            page.wait_for_selector("li[data-test='jobListing']", timeout=30000)
            page.wait_for_timeout(3000)
            
            # Step 5: Apply Remote filter and ensure others are unchecked
            try:
                # Uncheck Onsite/Hybrid if they are active
                page.evaluate('''() => {
                    const buttons = Array.from(document.querySelectorAll('button, [role="button"]'));
                    buttons.forEach(btn => {
                        const txt = (btn.textContent || '').trim().toLowerCase();
                        const checked = btn.getAttribute('aria-checked') === 'true' || 
                                        btn.getAttribute('data-checked') === 'true' || 
                                        btn.querySelector('input:checked') != null;
                        
                        if ((txt.includes('onsite') || txt.includes('on-site') || txt.includes('hybrid')) && checked) {
                            btn.click();
                        }
                    });
                }''')
                page.wait_for_timeout(1500)

                page.click("button[data-test='remoteWorkType']", timeout=10000)
                page.wait_for_timeout(3000)
                try:
                    page.wait_for_load_state("networkidle", timeout=10000)
                except TimeoutError:
                    pass
            except TimeoutError:
                print("  Remote button not found")

            if "remoteWorkType=1" not in page.url:
                print("  WARNING: the Remote filter did not apply (remoteWorkType=1 is missing from the address); stopping so no on-site jobs are saved.")
                return

            # Step 5b: Apply Glassdoor's own "Date posted: Last day" filter via
            # its fromAge=1 URL param, instead of loading every listing (most
            # of which are >24h old) and filtering them out client-side after
            # the fact - this cuts the number of jobs to load and process from
            # dozens/hundreds down to just the ones already within 24h.
            try:
                sep = "&" if "?" in page.url else "?"
                date_filtered_url = f"{page.url}{sep}fromAge=1"
                page.goto(date_filtered_url, timeout=30000, wait_until="domcontentloaded")
                page.wait_for_timeout(2000)
            except Exception as e:
                print(f"  Could not apply date filter, falling back to client-side filtering only: {e}")

            date_filter_on = "fromAge=1" in page.url
            if not date_filter_on:
                print("  WARNING: the 'last day' filter is not in the address; jobs without a readable age will be skipped.")

            # Step 6: Load all jobs by clicking "Show more jobs" or scrolling
            max_load_attempts = 50
            load_attempts = 0
            last_count = 0
            consecutive_no_increase = 0
            
            while load_attempts < max_load_attempts and consecutive_no_increase < 3:
                # Get current job count
                current_count = len(page.query_selector_all("li[data-test='jobListing']"))
                
                # Scroll to bottom to trigger lazy loading
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                page.wait_for_timeout(1500)
                
                # Try to find and click "Show more" button with multiple selectors
                clicked = False
                button_selectors = [
                    "button:has-text('Show more jobs')",
                    "button:has-text('Show More Jobs')",
                    "button:has-text('load more jobs')",
                    "button.button_Button__o_a9q.button-base_Button__zzUq2",
                    "[data-test='load-more-jobs']",
                    "button[aria-label*='more jobs']"
                ]
                
                for selector in button_selectors:
                    try:
                        btn = page.query_selector(selector)
                        if btn and btn.is_visible():
                            btn.click()
                            page.wait_for_timeout(3000)
                            clicked = True
                            break
                    except Exception:
                        continue

                # Check if job count increased
                new_count = len(page.query_selector_all("li[data-test='jobListing']"))
                if new_count > current_count:
                    consecutive_no_increase = 0
                else:
                    consecutive_no_increase += 1

                load_attempts += 1

            job_cards = page.query_selector_all("li[data-test='jobListing']")
            total = len(job_cards)
            print(f"Found {total} jobs")

            # Step 8: Process all jobs
            processed = 0
            for idx, card in enumerate(job_cards, 1):
                try:
                    # Skip Easy Apply
                    if card.query_selector("[data-test='easy-apply'], .easy-apply, div:has-text('Easy Apply')"):
                        print(f"  [{idx}] Skipped: Easy Apply (no employer link)")
                        continue

                    # Extract job age and filter (skip if older than 24 hours)
                    job_age = "N/A"
                    age_days = 0  # Default to accepting if age is missing, because fromAge=1 URL filter is applied
                    age_el = card.query_selector("[data-test='job-age'], .JobCard_listingAge__jJsuc, .jobAge, .listingAge")
                    if age_el:
                        job_age = age_el.inner_text().strip()
                        age_days = parse_job_age(job_age)

                    # The cards show no posting age; the last-day filter (fromAge=1) is on, so say so explicitly.
                    if job_age == "N/A" and date_filter_on:
                        job_age = "Within 24h"

                    if age_days > 1:
                        print(f"  [{idx}] Skipped: older than 24h ({job_age})")
                        continue

                    # Get job URL
                    link = card.query_selector("a[data-test='job-link'], a[href*='job/']")
                    if not link:
                        print(f"  [{idx}] Skipped: no job link on the card")
                        continue

                    job_url = link.get_attribute("href")
                    if not job_url.startswith("http"):
                        job_url = f"https://www.glassdoor.com{job_url}"

                    # Open job page
                    job_page = context.new_page()
                    job_page.goto(job_url, timeout=30000)
                    job_page.wait_for_selector("[data-test='job-title'], div.jobTitle, h1", timeout=15000)
                    job_page.wait_for_timeout(1500)
                    
                    # Extract data
                    company = "N/A"
                    for sel in [".heading_Heading__aomVx.heading_Subhead__jiUbT", "[data-test='employer-name']", "div.employerName", "span[data-test='employer-name']"]:
                        try:
                            el = job_page.wait_for_selector(sel, timeout=3000)
                            if el:
                                company = el.inner_text().strip()
                                break
                        except TimeoutError:
                            continue
                    
                    job_t = "N/A"
                    for sel in [".heading_Heading__aomVx.heading_Level1__w42c9", "[data-test='job-title']", "div.jobTitle", "h1"]:
                        try:
                            el = job_page.wait_for_selector(sel, timeout=3000)
                            if el:
                                job_t = el.inner_text().strip()
                                break
                        except TimeoutError:
                            continue
                    
                    loc = "N/A"
                    for sel in [".JobDetails_badgeStyle__xaoxT", "[data-test='job-location']", "div.location", "span[data-test='job-location']"]:
                        try:
                            el = job_page.wait_for_selector(sel, timeout=3000)
                            if el:
                                loc = el.inner_text().strip()
                                break
                        except TimeoutError:
                            continue
                    
                    description = read_description(job_page)

                    # Same job already in the sheet from another platform (or
                    # earlier in this run)? Skip before spending time on the
                    # Apply-click flow below.
                    job_key = (company.strip().lower(), job_t.strip().lower())
                    if job_key in seen_jobs:
                        print(f"  [{idx}] Skipped: already in the sheet ({company} / {job_t})")
                        job_page.close()
                        continue

                    # Click Apply
                    applied = False
                    for sel in [
                        "button:has-text('Apply on employer site')",
                        "[data-test='apply-button']",
                        "button:has-text('Apply Now')",
                        "a[data-test='apply-button']"
                    ]:
                        try:
                            el = job_page.wait_for_selector(sel, timeout=4000)
                            if el and el.is_visible():
                                applied = True
                                job_page = _click_and_follow(context, job_page, el)
                                break
                        except TimeoutError:
                            continue

                    if not applied:
                        print(f"  [{idx}] Skipped: no 'Apply on employer site' button")
                        job_page.close()
                        continue

                    # Click Start Application if present
                    try:
                        el = job_page.wait_for_selector(
                            "button:has-text('Start my application'), [data-test='start-application-button'], button:has-text('Start application')",
                            timeout=3000)
                        if el and el.is_visible():
                            job_page = _click_and_follow(context, job_page, el)
                    except TimeoutError:
                        pass

                    app_url = job_page.url.strip()
                    if not app_url or app_url == "about:blank":
                        print(f"  [{idx}] Skipped: application tab never left about:blank (not a ready-to-apply link)")
                        job_page.close()
                        continue

                    # "Apply on employer site" on most Glassdoor postings routes
                    # through Indeed's account-linking sign-in flow, and Indeed's
                    # own Cloudflare bot-check blocks the automated browser there
                    # almost every time - not something to try to defeat. Only
                    # links that land somewhere real should be saved, so skip
                    # the job entirely here rather than saving a link the
                    # applicant still can't actually apply from.
                    try:
                        page_text = job_page.title() + " " + job_page.inner_text("body")[:500]
                    except Exception:
                        page_text = ""
                    if looks_like_captcha(app_url, page_text):
                        print(f"  [{idx}] Skipped: CAPTCHA/verification page (not a ready-to-apply link)")
                        job_page.close()
                        continue

                    # "Apply on employer site" sometimes just opens Glassdoor's
                    # own apply-tracking/interstitial page rather than truly
                    # leaving Glassdoor - not a page the applicant can actually
                    # submit an application from, so it doesn't count as
                    # ready-to-apply either.
                    if "glassdoor.com" in app_url:
                        print(f"  [{idx}] Skipped: still on Glassdoor, never reached the employer's application page ({app_url[:120]})")
                        job_page.close()
                        continue

                    # Duplicate check: skip only if this exact application link was
                    # already added (in this run or a past one). Multiple postings from
                    # the same company are fine as long as each has its own link.
                    if app_url in seen_links:
                        print(f"  [{idx}] Skipped: this application link is already in the sheet")
                        job_page.close()
                        continue

                    saved = True
                    if sheet:
                        saved = append_job(sheet, {
                            "company": company, "title": job_t, "location": loc, "age": job_age,
                            "link": app_url, "platform": "Glassdoor", "description": description,
                        }, seen_links, seen_jobs)
                    if saved:
                        ws.append([company, job_t, loc, job_age, app_url])
                        processed += 1
                    job_page.close()
                    
                    # Test mode limit
                    if TEST_LIMIT and processed >= TEST_LIMIT:
                        print(f"  Test limit reached ({TEST_LIMIT} jobs). Stopping early.")
                        break
                    
                except Exception as e:
                    print(f"  [{idx}] Error: {e}")
                    # Close job_page plus any other stray tab this job may have
                    # opened (e.g. a popup from an unexpected click) that
                    # job_page never got reassigned to before the error hit -
                    # otherwise it's left open (often on a CAPTCHA/verification
                    # page) for the rest of the run.
                    for p in list(context.pages):
                        if p is not page:
                            try:
                                p.close()
                            except Exception:
                                pass
                    continue
            
            # Save Excel
            wb.save("glassdoor.xlsx")
            print(f"\nDone! Saved {processed} jobs to glassdoor.xlsx")
            
        except Exception as e:
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()
        finally:
            context.close()

if __name__ == "__main__":
    main()
