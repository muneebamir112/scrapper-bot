from patchright.sync_api import sync_playwright, TimeoutError
from openpyxl import Workbook
import gspread
from google.oauth2.service_account import Credentials
import os
import re
import sys
from datetime import date

# Persistent Chrome profile used by the stealth browser. Login state (and any
# cookies needed to avoid Glassdoor's bot checks) is kept here between runs.
PROFILE_DIR = r'C:\Users\webNcodes\AppData\Local\Google\Chrome\User Data\Profile 9'
# PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chrome_profile")

# Google Sheet that scraped jobs are mirrored into (in addition to glassdoor.xlsx).
# Existing sheet layout: No | Company Name | Job Title | Location | Job Age | Job Link | Date Added | Source
import os
from dotenv import load_dotenv
load_dotenv(r"c:\Users\webNcodes\Desktop\webncodes\Job-Bot\.env")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1Kva2y5-54LXBWMTzNk_xp524ZE7N-CWiqL3VGATUZVM")
SERVICE_ACCOUNT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "service_account.json")
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
        worksheet = client.open_by_key(GOOGLE_SHEET_ID).sheet1
        return worksheet
    except Exception as e:
        print(f"  Google Sheets connection failed: {e}")
        return None

def find_next_sheet_slot(worksheet):
    """The sheet already has a No/Company Name/.../Status/Date Added layout with
    pre-filled blank placeholder rows (Status='Pending' set ahead of time). Find the
    first row whose Company Name cell is still empty so new jobs fill those rows in
    order instead of being appended after them, and derive the next 'No' value from
    the row directly above it."""
    company_col = worksheet.col_values(2)  # column B, includes header at index 0
    next_row = len(company_col) + 1
    for i in range(1, len(company_col)):
        if not company_col[i].strip():
            next_row = i + 1
            break
    return next_row, next_row - 1

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
    age_text = age_text.lower().strip()
    # Match patterns: 23d, 5h, 2w, 1mo, 3mos, etc.
    match = re.match(r'(\d+)(mo|mos|[dhmsw])', age_text)
    if not match:
        return 0
    value = int(match.group(1))
    unit = match.group(2)
    if unit in ('h',):
        return value / 24
    elif unit in ('d',):
        return value
    elif unit in ('w',):
        return value * 7
    elif unit in ('m', 'mo', 'mos'):
        return value * 30
    elif unit == 's':
        return value / (24 * 7 * 4)
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

def _wait_for_real_navigation(page, timeout_ms=8000):
    """A newly opened tab sits on about:blank for a moment until its actual
    navigation lands - waiting a fixed delay and then reading page.url risks
    reading it while it's still blank (which was ending up as the captured
    application link). Wait for the URL to actually change instead."""
    try:
        page.wait_for_url(lambda url: url != "about:blank", timeout=timeout_ms)
    except TimeoutError:
        pass
    try:
        page.wait_for_load_state("domcontentloaded", timeout=2000)
    except Exception:
        pass


def _switch_to_new_tab(context, current_page):
    """If a click just opened a new tab (e.g. an employer/Indeed apply flow
    breaking out of Glassdoor's own job-detail tab), close the now-stale
    current_page and return the new tab; otherwise return current_page
    unchanged. Without this, the old tab - often left sitting on an Indeed
    CAPTCHA/verification page - stays open for the rest of the run."""
    pages = context.pages
    if len(pages) > 1 and pages[-1] != current_page:
        new_page = pages[-1]
        try:
            current_page.close()
        except Exception:
            pass
        _wait_for_real_navigation(new_page)
        return new_page
    return current_page


def make_link_formula(url):
    """Wrap a URL in a HYPERLINK() formula so it renders as a clickable link
    in the sheet (gspread's default FORMATTED_VALUE reads still return the
    plain URL text, so nothing downstream that reads the sheet is affected)."""
    escaped = url.replace('"', '""')
    return f'=HYPERLINK("{escaped}", "{escaped}")'

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
    sheet_next_row, sheet_next_no = find_next_sheet_slot(sheet) if sheet else (None, None)
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
            page.goto("https://www.glassdoor.com/Job/index.htm", timeout=30000, wait_until="domcontentloaded")
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
            
            # Step 5: Apply Remote filter
            try:
                page.click("button[data-test='remoteWorkType']", timeout=10000)
                page.wait_for_timeout(3000)
                try:
                    page.wait_for_load_state("networkidle", timeout=10000)
                except TimeoutError:
                    pass
            except TimeoutError:
                print("  Remote button not found")

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
                        continue

                    # Extract job age and filter (skip if older than 24 hours)
                    job_age = "N/A"
                    age_el = card.query_selector("[data-test='job-age'], .JobCard_listingAge__jJsuc, .jobAge, .listingAge")
                    if age_el:
                        job_age = age_el.inner_text().strip()
                        age_days = parse_job_age(job_age)
                        if age_days > 1:
                            continue

                    # Get job URL
                    link = card.query_selector("a[data-test='job-link'], a[href*='job/']")
                    if not link:
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
                    
                    # Same job already in the sheet from another platform (or
                    # earlier in this run)? Skip before spending time on the
                    # Apply-click flow below.
                    job_key = (company.strip().lower(), job_t.strip().lower())
                    if job_key in seen_jobs:
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
                                el.click()
                                applied = True
                                job_page.wait_for_timeout(3000)
                                # This click sometimes opens the employer's/Indeed's
                                # apply page in a new tab rather than navigating
                                # job_page itself - switch to it and close the old
                                # tab so it doesn't sit open (often on an Indeed
                                # CAPTCHA/verification page) for the rest of the run.
                                job_page = _switch_to_new_tab(context, job_page)
                                break
                        except TimeoutError:
                            continue

                    if not applied:
                        job_page.close()
                        continue

                    # Click Start Application if present
                    for sel in [
                        "button:has-text('Start my application')",
                        "[data-test='start-application-button']",
                        "button:has-text('Start application')"
                    ]:
                        try:
                            el = job_page.wait_for_selector(sel, timeout=3000)
                            if el and el.is_visible():
                                el.click()
                                job_page.wait_for_timeout(10000)
                                job_page = _switch_to_new_tab(context, job_page)
                                break
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
                        print(f"  [{idx}] Skipped: still on Glassdoor, never reached the employer's application page")
                        job_page.close()
                        continue

                    # Duplicate check: skip only if this exact application link was
                    # already added (in this run or a past one). Multiple postings from
                    # the same company are fine as long as each has its own link.
                    if app_url in seen_links:
                        job_page.close()
                        continue

                    ws.append([company, job_t, loc, job_age, app_url])
                    seen_links.add(app_url)
                    seen_jobs.add(job_key)

                    if sheet:
                        try:
                            sheet_row = [
                                str(sheet_next_no), company, job_t, loc, job_age,
                                make_link_formula(app_url), date.today().strftime("%m/%d/%Y"),
                                "Glassdoor",
                            ]
                            sheet.update(range_name=f"A{sheet_next_row}:H{sheet_next_row}", values=[sheet_row], value_input_option="USER_ENTERED")
                            sheet_next_row += 1
                            sheet_next_no += 1
                        except Exception as e:
                            print(f"  [{idx}] Google Sheets sync failed: {e}")

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
