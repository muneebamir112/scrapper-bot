import asyncio
import sys
from urllib.parse import quote
import pandas as pd
from patchright.async_api import async_playwright
import re
import os
import gspread
from google.oauth2.service_account import Credentials
from datetime import date

# Persistent Chrome profile for the stealth browser (same patchright approach
# GlassD's scraper uses), kept separate from GlassD's/Hiring Cafe's own
# profiles so each platform's login/session state stays independent.
PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chrome_profile")

# Same Google Sheet + service account GlassD's scraper uses, so job links from
# every platform land in one place (Sheet1: No | Company Name | Job Title |
# Location | Job Age | Job Link | Date Added | Source).
GOOGLE_SHEET_ID = "1FsPR9t-BB1GZ6kWfANnrDfq4p9XobDuA1tVB4D2sJDg"
SERVICE_ACCOUNT_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "GlassD", "service_account.json"
)


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
        worksheet = client.open_by_key(GOOGLE_SHEET_ID).sheet1
        return worksheet
    except Exception as e:
        print(f"  Google Sheets connection failed: {e}")
        return None


def find_next_sheet_slot(worksheet):
    """Find the first row whose Company Name cell (column B) is still empty so
    new jobs fill pre-existing placeholder rows in order, deriving the next
    'No' value from the row directly above it."""
    company_col = worksheet.col_values(2)
    next_row = len(company_col) + 1
    for i in range(1, len(company_col)):
        if not company_col[i].strip():
            next_row = i + 1
            break
    return next_row, next_row - 1


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
    link = (job.get("Apply Link") or "").strip()
    if not link or link in seen_links:
        return next_row, next_no
    company = (job.get("Company") or "").strip().lower()
    title = (job.get("Job Title") or "").strip().lower()
    job_key = (company, title)
    if company and title and job_key in seen_jobs:
        return next_row, next_no
    try:
        sheet_row = [
            str(next_no), job.get("Company", ""), job.get("Job Title", ""),
            job.get("Location", ""), job.get("Job Age", ""), make_link_formula(link),
            date.today().strftime("%m/%d/%Y"), "Jobgether",
        ]
        sheet.update(range_name=f"A{next_row}:H{next_row}", values=[sheet_row], value_input_option="USER_ENTERED")
        seen_links.add(link)
        seen_jobs.add(job_key)
        return next_row + 1, next_no + 1
    except Exception as e:
        print(f"  Google Sheets sync failed for '{job.get('Job Title', '')}': {e}")
        return next_row, next_no


def slugify(text):
    """Convert a job title/keyword into the URL slug jobgether.com's
    remote-jobs pages expect (e.g. 'Python Developer' -> 'python-developer')."""
    text = text.strip().lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def is_within_24h(age_text):
    """Best-effort check of a 'X ago' style string. 'Today'/hour/minute
    granularity is always within 24h; a day count of 1 or more (including
    '30+ days ago'), or any week/month unit, means more than 24 hours have
    already passed."""
    if not age_text:
        return True
    text = age_text.lower().strip()
    if "today" in text or "just" in text or "now" in text:
        return True
    match = re.search(r'(\d+)\+?\s*(day|days|hour|hours|minute|minutes|week|weeks|month|months)', text)
    if not match:
        return True
    value = int(match.group(1))
    unit = match.group(2)
    if "hour" in unit or "minute" in unit:
        return True
    if "day" in unit:
        return value < 1
    return False  # week/month


async def main():
    # jobgether.com now redirects the old /remote-jobs/<location>/<keyword>
    # slug URL to a plain /search-offers?location=<location> page with the
    # keyword silently dropped, returning every job instead of a filtered
    # set. The keyword must instead be passed as its own query param.
    if len(sys.argv) > 1:
        keyword = sys.argv[1]
    else:
        try:
            keyword = input("Enter job keyword (e.g., 'full stack engineer'): ").strip()
        except EOFError:
            keyword = "integration engineer"  # Default from question tool

    location_slug = "united-states"
    # sort=date orders cards newest-first (jobgether's default is relevance,
    # which mixes fresh and month-old postings together) so the loop below
    # can stop as soon as it hits the first card older than 24h instead of
    # scanning every card on the page to find the few recent ones.
    #
    # includeHybrid=false pins the site's own "Remote type" filter to
    # remote-only. jobgether.com already defaults to remote-only (its filter
    # panel says "Only full remote jobs are displayed by default"), but this
    # scraper reuses a persistent Chrome profile across runs, so relying on
    # that default would silently break if any past session ever checked
    # the "Include hybrid jobs" box - that state could persist. Setting it
    # explicitly in the URL guarantees remote-only regardless.
    target_url = (
        f"https://jobgether.com/search-offers?location={location_slug}"
        f"&keyword={quote(keyword)}&sort=date&includeHybrid=false"
    )

    results = []

    async with async_playwright() as p:
        # Launch our own stealth-patched Chrome (same patchright approach
        # GlassD's scraper uses) instead of depending on an already-running
        # Chrome with a remote debugging port.
        os.makedirs(PROFILE_DIR, exist_ok=True)
        context = await p.chromium.launch_persistent_context(
            PROFILE_DIR,
            channel="chrome",
            headless=False,
            no_viewport=True,
        )

        if len(context.pages) > 0:
            page = context.pages[0]
        else:
            page = await context.new_page()

        try:
            await page.goto(target_url, wait_until="domcontentloaded", timeout=60000)
            await page.wait_for_timeout(4000)
        except Exception as e:
            print(f"Navigation error: {e}")

        try:
            cards = await page.evaluate("""() => {
                const offerLinks = Array.from(document.querySelectorAll('a[href^="/offer/"]'));
                const seen = new Set();
                const out = [];
                for (const a of offerLinks) {
                    const href = a.getAttribute('href');
                    if (seen.has(href) || href === '/offer/undefined') continue;
                    seen.add(href);
                    let posted = '';
                    const parent = a.closest('div');
                    if (parent) {
                        const sib = parent.querySelector('div.text-xs.text-gray-500, div[class*="text-gray-500"]');
                        if (sib) posted = sib.textContent.trim();
                    }
                    let company = '';
                    const card = a.closest('div.p-6') || a.closest('[class*="p-6"]');
                    if (card) {
                        const companyA = card.querySelector('a[href*="/remote-jobs/company-"], a[href*="/search-offers?company="]');
                        if (companyA) company = companyA.textContent.trim();
                    }
                    out.push({title: a.textContent.trim(), href: href, posted: posted, company: company});
                }
                return out;
            }""")
            print(f"Found {len(cards)} job cards")

            # Cards are sorted newest-first (sort=date), so the first one
            # older than 24h means everything after it is too - no need to
            # keep scanning the rest of the page.
            fresh_cards = []
            for c in cards:
                if is_within_24h(c["posted"]):
                    fresh_cards.append(c)
                else:
                    break
            print(f"{len(fresh_cards)} job(s) posted within the last 24 hours")

            # Connect once, before extraction starts, so each job below can
            # be written to the sheet as soon as it's scraped rather than
            # waiting to batch-sync the whole list at the end.
            sheet = get_google_sheet()
            sheet_next_row, sheet_next_no = find_next_sheet_slot(sheet) if sheet else (None, None)
            seen_links = load_seen_links(sheet) if sheet else set()
            seen_jobs = load_seen_jobs(sheet) if sheet else set()
            print(f"  Loaded {len(seen_links)} existing job links and {len(seen_jobs)} company/title pairs for duplicate check")

            for i, card in enumerate(fresh_cards, start=1):
                try:
                    full_url = "https://jobgether.com" + card["href"]
                    job_page = await context.new_page()
                    await job_page.goto(full_url, wait_until="domcontentloaded", timeout=30000)
                    await job_page.wait_for_timeout(2000)

                    # The real external application link is a plain <a href>
                    # already in the page (data-role="apply-guest" for
                    # anonymous visitors, "apply-logged" once signed in) -
                    # no click/new-tab dance needed.
                    apply_link = await job_page.evaluate("""() => {
                        const a = document.querySelector('a[data-role="apply-guest"]')
                            || document.querySelector('a[data-role="apply-logged"]');
                        return a ? a.getAttribute('href') : null;
                    }""")

                    if apply_link and looks_like_captcha(apply_link, ""):
                        apply_link = None

                    if apply_link:
                        job_data = {
                            "Job Title": card["title"],
                            "Company": card["company"],
                            "Job Age": card["posted"],
                            "Location": "United States",
                            "Apply Link": apply_link,
                        }
                        results.append(job_data)
                        if sheet:
                            sheet_next_row, sheet_next_no = sync_one_job_to_sheet(
                                sheet, job_data, seen_links, seen_jobs, sheet_next_row, sheet_next_no
                            )

                    await job_page.close()
                    await page.wait_for_timeout(300)
                except Exception as e:
                    print(f"  Error processing job: {e}")
                    continue

            # Save to Excel - one fixed filename overwritten each run (same
            # as Glassdoor's glassdoor.xlsx and Hiring Cafe's jobs.xlsx),
            # not a new timestamped file every run - this used to pile up
            # one file per keyword per run with no cleanup, since the Google
            # Sheet above is the actual persistent destination and this is
            # just a local backup snapshot of the latest run.
            df = pd.DataFrame(results)
            excel_file = "jobgether.xlsx"
            df.to_excel(excel_file, index=False)
            print(f"\nSaved {len(results)} jobs to {excel_file}")

        finally:
            await context.close()

if __name__ == "__main__":
    asyncio.run(main())