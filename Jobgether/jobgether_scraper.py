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

if getattr(sys, 'frozen', False):
    CURRENT_DIR = os.path.dirname(sys.executable)
    LAUNCHER_DIR = CURRENT_DIR
else:
    CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = os.path.dirname(os.path.dirname(CURRENT_DIR))
    LAUNCHER_DIR = os.path.join(BASE_DIR, "Job-bot-launcher")

# Persistent Chrome profile for the stealth browser (same patchright approach
# GlassD's scraper uses), kept separate from GlassD's/Hiring Cafe's own
# profiles so each platform's login/session state stays independent.
PROFILE_DIR = os.path.join(CURRENT_DIR, "profiles", "jobgether_profile")

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


import html as _html
import json
import time
from urllib.parse import urlparse


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


async def read_description(page):
    """Description text of the page that is open: JSON-LD first, then the main text block."""
    candidates = []
    try:
        candidates.append(description_from_jsonld(await page.evaluate(JSONLD_JS)))
    except Exception:
        pass
    try:
        candidates.append(await page.evaluate(BLOCK_JS))
    except Exception:
        pass
    return pick_description(candidates)


def slugify(text):
    """Convert a job title/keyword into the URL slug jobgether.com's
    remote-jobs pages expect (e.g. 'Python Developer' -> 'python-developer')."""
    text = text.strip().lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def is_within_24h(age_text):
    """Check a 'posted X ago' text. Hours, minutes, 'today' and 'just now' count. Jobgether only
    reports whole days, so '1 day ago' (its freshest day label) is accepted as well; 2+ days,
    weeks, months and unrecognised text are not."""
    text = str(age_text or "").lower().strip()
    if not text:
        return False
    if "just now" in text or "today" in text or text == "now":
        return True
    match = re.search(r"(\d+)\+?\s*(d|days?|h|hrs?|hours?|m|mins?|minutes?)\b", text)
    if not match:
        return False
    unit, value = match.group(2), int(match.group(1))
    if unit.startswith("d"):
        return value <= 1
    return unit.startswith("m") or value <= 24


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
        f"&keyword={quote(keyword)}&sort=date&includeHybrid=false&includeOnsite=false"
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

        for attempt in range(4):
            try:
                print("  Clearing local storage to prevent default saved modes...")
                await page.goto("https://jobgether.com/", timeout=60000)
                await page.evaluate("window.localStorage.clear(); window.sessionStorage.clear();")
                await page.wait_for_timeout(1000)

                await page.goto(target_url, wait_until="domcontentloaded", timeout=60000)
                await page.wait_for_timeout(4000)
                break
            except Exception as e:
                if attempt == 3:
                    print("Network error: Could not reach jobgether.com after multiple attempts. Please check your internet connection.")
                    return
                print(f"Network hiccup reaching jobgether.com, retrying ({attempt + 1}/3)...")
                await asyncio.sleep(2)

        try:
            cards = await page.evaluate(r"""() => {
                const offerLinks = Array.from(document.querySelectorAll('a[href^="/offer/"]'));
                const seen = new Set();
                const out = [];
                for (const a of offerLinks) {
                    const href = a.getAttribute('href');
                    if (seen.has(href) || href === '/offer/undefined') continue;
                    seen.add(href);
                    let posted = '';
                    const postedCard = a.closest('div.group') || a.parentElement.parentElement.parentElement;
                    if (postedCard) {
                        const fullText = postedCard.innerText;
                        const lower = fullText.toLowerCase();
                        if (lower.includes("just now")) {
                            posted = "Just Now";
                        } else if (lower.includes("today")) {
                            posted = "Today";
                        } else {
                            const match = lower.match(/(\d+\+?\s*(?:day|days|hour|hours|minute|minutes|week|weeks|month|months)(?:\s*ago)?)/i);
                            if (match) {
                                posted = match[1];
                            } else {
                                posted = fullText.split('\n')[1] || ''; // Fallback
                            }
                        }
                    }
                    let company = '';
                    const card = a.closest('li') || a.closest('div.group') || a.parentElement.parentElement.parentElement;
                    if (card) {
                        const companyA = card.querySelector('a[href*="/remote-jobs/company-"], a[href*="/search-offers?company="]');
                        if (companyA) company = companyA.textContent.trim();
                        // fallback to finding the img alt attribute just in case the link goes away
                        if (!company) {
                            const img = card.querySelector('img[alt]');
                            if (img) company = img.getAttribute('alt');
                        }
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
            if sheet:
                ensure_header(sheet)
            seen_links = load_seen_links(sheet) if sheet else set()
            seen_jobs = load_seen_jobs(sheet) if sheet else set()
            print(f"  Loaded {len(seen_links)} existing job links and {len(seen_jobs)} company/title pairs for duplicate check")

            sem = asyncio.Semaphore(2)

            async def process(i, card):
                async with sem:
                    # Already in the sheet (same company + title)? Skip before opening the offer page.
                    key = (card["company"].strip().lower(), card["title"].strip().lower())
                    if key[0] and key[1] and key in seen_jobs:
                        return
                    job_page = None
                    try:
                        full_url = "https://jobgether.com" + card["href"]
                        job_page = await context.new_page()
                        await job_page.goto(full_url, wait_until="domcontentloaded", timeout=30000)
                        try:
                            await job_page.wait_for_selector('a[data-role="apply-guest"], a[data-role="apply-logged"]', timeout=8000)
                        except Exception:
                            pass  # the check below reports "could not find apply link"

                        # The real external application link is a plain <a href>
                        # already in the page (data-role="apply-guest" for
                        # anonymous visitors, "apply-logged" once signed in) -
                        # no click/new-tab dance needed.
                        apply_link = await job_page.evaluate("""() => {
                            const a = document.querySelector('a[data-role="apply-guest"]')
                                || document.querySelector('a[data-role="apply-logged"]');
                            return a ? a.getAttribute('href') : null;
                        }""")

                        description = await read_description(job_page) if apply_link else ""

                        if apply_link and looks_like_captcha(apply_link, ""):
                            print(f"  [{i}] Skipped: apply link looks like a CAPTCHA gate")
                            apply_link = None
                        elif not apply_link:
                            print(f"  [{i}] Skipped: could not find apply link on page")
                        else:
                            print(f"  [{i}] Saving job: {apply_link}")

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
                                append_job(sheet, {
                                    "company": card["company"], "title": card["title"], "location": "United States",
                                    "age": card["posted"], "link": apply_link, "platform": "Jobgether",
                                    "description": description,
                                }, seen_links, seen_jobs)
                    except Exception as e:
                        print(f"  Error processing job: {e}")
                    finally:
                        if job_page is not None:
                            try:
                                await job_page.close()
                            except Exception:
                                pass

            print(f"  Processing {len(fresh_cards)} job(s), 2 at a time...")
            await asyncio.gather(*(process(i, c) for i, c in enumerate(fresh_cards, start=1)), return_exceptions=True)

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