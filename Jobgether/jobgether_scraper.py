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
import time
from urllib.parse import quote, urlparse
import pandas as pd
from patchright.async_api import async_playwright
import re
import os
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date

PLATFORM_NAME = "Jobgether"
_T0 = time.time()
STATS = {"cards": 0, "fresh": 0, "dup": 0, "skip": 0, "saved": 0, "bad": 0, "status": "ok"}


def log(level, msg):
    """Standard log line: YYYY-MM-DD HH:MM:SS [PLATFORM] [LEVEL] Message"""
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{now_str} [{PLATFORM_NAME}] [{level:<5}] {msg}", flush=True)


def short(url):
    """Host + path only, cut to 70 characters. No query string: links can carry tokens."""
    u = urlparse(str(url))
    return (u.netloc + u.path)[:70]


def jid(link):
    """Short id of a job link, to follow one job through the log."""
    return urlparse(str(link)).path.rsplit("/", 1)[-1][:12]


def summary():
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    elapsed = int(time.time() - _T0)
    print(f"\n{'='*70}\n{now_str} [{PLATFORM_NAME}] Run Completed (Elapsed: {elapsed}s)\n"
          f"Status: {STATS['status']} | Scanned: {STATS['cards']} | Fresh (24h): {STATS['fresh']} | "
          f"Saved: {STATS['saved']} | Duplicates: {STATS['dup']} | Skipped: {STATS['skip']} | Failed: {STATS['bad']}\n"
          f"{'='*70}\n", flush=True)

if getattr(sys, 'frozen', False):
    CURRENT_DIR = os.path.dirname(sys.executable)
    LAUNCHER_DIR = CURRENT_DIR
else:
    CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = os.path.dirname(os.path.dirname(CURRENT_DIR))
    LAUNCHER_DIR = os.path.join(BASE_DIR, "Job-bot-launcher")

if LAUNCHER_DIR not in sys.path:
    sys.path.insert(0, LAUNCHER_DIR)

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
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "").strip()
SERVICE_ACCOUNT_FILE = os.path.join(LAUNCHER_DIR, "service_account.json")


def get_google_sheet():
    """Connect to Sheet1 of the shared Google Sheet. Returns None (and prints why)
    if the service account file is missing or the connection fails, so the
    scraper can keep running and just skip the Sheets sync."""
    if not GOOGLE_SHEET_ID:
        log("WARN", "Google Sheets sync disabled: GOOGLE_SHEET_ID not set in .env")
        return None
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        log("WARN", "Google Sheets sync disabled: service account file not found")
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
        log("WARN", f"Google Sheets connection failed: {e}")
        return None


def ensure_header(worksheet):
    """Create the header row if the tab is empty (columns J & K hold question docs)."""
    company_col = worksheet.col_values(2)
    if len(company_col) == 0 or (len(company_col) == 1 and not company_col[0].strip()):
        headers = ["No.", "Company Name", "Job Title", "Location", "Job Age", "Job Link", "Date Posted", "Platform", "Job Description", "Questions Doc Link", "Questions Count"]
        worksheet.update(range_name="A1:K1", values=[headers], value_input_option="USER_ENTERED")
        try:
            worksheet.format("A1:K1", {
                "backgroundColor": {"red": 0.99, "green": 0.89, "blue": 0.58},
                "textFormat": {"bold": True, "fontSize": 10},
                "horizontalAlignment": "CENTER",
            })
            worksheet.format("I2:I", {"numberFormat": {"type": "TEXT"}, "wrapStrategy": "CLIP"})
        except Exception as e:
            log("WARN", f"could not format headers: {e}")
    else:
        try:
            existing_headers = worksheet.row_values(1)
            if len(existing_headers) < 10:
                worksheet.update(range_name="J1:K1", values=[["Questions Doc Link", "Questions Count"]], value_input_option="USER_ENTERED")
                worksheet.format("J1:K1", {
                    "backgroundColor": {"red": 0.99, "green": 0.89, "blue": 0.58},
                    "textFormat": {"bold": True, "fontSize": 10},
                    "horizontalAlignment": "CENTER",
                })
        except Exception:
            pass


def load_seen_links(worksheet):
    """Application links already logged in the sheet ('Job Link' column),
    used to skip a job only if this exact link was already added before."""
    try:
        headers = [h.strip().lower() for h in worksheet.row_values(1)]
        col_idx = headers.index("job link") + 1 if "job link" in headers else 6
        link_col = worksheet.col_values(col_idx)
    except Exception as e:
        log("WARN", f"could not read existing job links from the sheet: {e}")
        return set()
    return {v.strip() for v in link_col[1:] if v.strip()}


def load_seen_jobs(worksheet):
    """(Company, Job Title) pairs already in the sheet,
    normalized for case/whitespace. Finds columns dynamically by header names."""
    try:
        headers = [h.strip().lower() for h in worksheet.row_values(1)]
        c_idx = headers.index("company name") + 1 if "company name" in headers else 2
        t_idx = headers.index("job title") + 1 if "job title" in headers else 3
        company_col = worksheet.col_values(c_idx)
        title_col = worksheet.col_values(t_idx)
    except Exception as e:
        log("WARN", f"could not read existing company/title pairs from the sheet: {e}")
        return set()
    seen = set()
    for company, title in zip(company_col[1:], title_col[1:]):
        company = company.strip().lower()
        title = title.strip().lower()
        if company and title:
            seen.add((company, title))
    return seen


OFFLINE_EXCEL = "jobgether_offline.xlsx"


def save_offline_job(job):
    """Saves a job locally to jobgether_offline.xlsx only when Google Sheets is unreachable."""
    try:
        import openpyxl
        if os.path.exists(OFFLINE_EXCEL):
            wb = openpyxl.load_workbook(OFFLINE_EXCEL)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "OfflineJobs"
            ws.append(["company", "title", "location", "age", "link", "platform", "description"])
        
        ws.append([
            job.get("company", ""), job.get("title", ""), job.get("location", ""),
            job.get("age", ""), job.get("link", ""), job.get("platform", "Jobgether"),
            job.get("description", "")
        ])
        wb.save(OFFLINE_EXCEL)
        log("INFO", f"[Offline Cache] Saved {job.get('company')} | {job.get('title')} to {OFFLINE_EXCEL}")
    except Exception as e:
        log("ERR", f"[Offline Cache Error] Could not save to {OFFLINE_EXCEL}: {e}")


def sync_offline_jobs(sheet, seen_links, seen_jobs):
    """Syncs any locally cached offline jobs to Google Sheets once connection is available."""
    if not sheet or not os.path.exists(OFFLINE_EXCEL):
        return
    try:
        import openpyxl
        wb = openpyxl.load_workbook(OFFLINE_EXCEL)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) <= 1:
            try:
                os.remove(OFFLINE_EXCEL)
            except Exception:
                pass
            return

        headers = [str(h).strip().lower() for h in rows[0]]
        h_map = {h: i for i, h in enumerate(headers)}
        synced_count = 0
        remaining_rows = [rows[0]]

        for r in rows[1:]:
            company = r[h_map.get("company", 0)] or ""
            title = r[h_map.get("title", 1)] or ""
            location = r[h_map.get("location", 2)] or ""
            age = r[h_map.get("age", 3)] or ""
            link = r[h_map.get("link", 4)] or ""
            platform = r[h_map.get("platform", 5)] or "Jobgether"
            description = r[h_map.get("description", 6)] or ""

            job_dict = {
                "company": str(company), "title": str(title), "location": str(location),
                "age": str(age), "link": str(link), "platform": str(platform),
                "description": str(description)
            }
            if append_job(sheet, job_dict, seen_links, seen_jobs, is_offline_sync=True):
                synced_count += 1
            else:
                remaining_rows.append(r)

        if synced_count > 0:
            log("INFO", f"[Offline Sync] Successfully synced {synced_count} offline job(s) to Google Sheets!")

        if len(remaining_rows) <= 1:
            wb.close()
            try:
                os.remove(OFFLINE_EXCEL)
            except Exception:
                pass
        else:
            new_wb = openpyxl.Workbook()
            new_ws = new_wb.active
            new_ws.title = "OfflineJobs"
            for r in remaining_rows:
                new_ws.append(list(r))
            new_wb.save(OFFLINE_EXCEL)
    except Exception as e:
        log("WARN", f"[Offline Sync Warning] Error while syncing {OFFLINE_EXCEL}: {e}")


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


COOKIE_PATTERNS = [
    re.compile(r"we use cookies", re.IGNORECASE),
    re.compile(r"cookie policy", re.IGNORECASE),
    re.compile(r"cookie preferences", re.IGNORECASE),
    re.compile(r"accept all cookies", re.IGNORECASE),
    re.compile(r"by clicking [\"']?accept[\"']?", re.IGNORECASE),
    re.compile(r"privacy policy.*cookies", re.IGNORECASE),
    re.compile(r"manage preferences", re.IGNORECASE),
]


def looks_like_cookie_banner(text):
    """Returns True if the text is predominantly a cookie/GDPR consent notice."""
    t = text.strip()
    if len(t) < 1500:
        matches = sum(1 for p in COOKIE_PATTERNS if p.search(t))
        if matches >= 2:
            return True
        if len(t) < 600 and ("cookie" in t.lower() or "consent" in t.lower()):
            return True
    return False


def pick_description(candidates):
    """First candidate with at least 400 characters, else the longest non-empty one.
    Ignores candidates that are just cookie banners or privacy consent popups."""
    texts = [t for t in (clean_description(c) for c in candidates if c) if t and not looks_like_cookie_banner(t)]
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


def append_job(sheet, job, seen_links, seen_jobs, is_offline_sync=False):
    """Append one job to the sheet. Google picks the next free row, so scrapers running
    at the same time cannot overwrite each other. Skips duplicates and incomplete jobs.
    Returns True if the row was written."""
    link = str(job.get("link") or "").strip()
    key = (str(job.get("company") or "").strip().lower(), str(job.get("title") or "").strip().lower())
    who = f"{job.get('company', '')} | {job.get('title', '')}"
    if not link:
        STATS["skip"] += 1
        log("SKIP", f"{who}: no apply link")
        return False
    if link in seen_links or (key[0] and key[1] and key in seen_jobs):
        STATS["dup"] += 1
        log("DUP", f"{who}: {'link' if link in seen_links else 'company+title'} already in the sheet (at write)")
        return False
    problem = row_problem(job)
    if problem:
        STATS["skip"] += 1
        log("SKIP", f"{who}: {problem}")
        return False
    questions_doc = job.get("questions_doc", "")
    questions_count = job.get("questions_count", "None")
    doc_cell = make_link_formula(questions_doc) if questions_doc else ""

    # Dynamically map values to header columns so extra or moved columns never break alignment
    try:
        headers = [h.strip().lower() for h in sheet.row_values(1)]
    except Exception:
        headers = []

    if headers:
        col_map = {h: i for i, h in enumerate(headers)}
        row = [""] * len(headers)
        
        def set_val(name, val):
            if name in col_map:
                row[col_map[name]] = val

        set_val("no.", "=ROW()-1")
        set_val("company name", job["company"])
        set_val("job title", job["title"])
        set_val("location", job["location"])
        set_val("job age", job["age"])
        set_val("job link", make_link_formula(link))
        set_val("date posted", date.today().strftime("%m/%d/%Y"))
        set_val("platform", job["platform"])
        set_val("job description", description)
        set_val("questions doc link", doc_cell)
        set_val("questions count", questions_count)
    else:
        row = [
            "=ROW()-1", job["company"], job["title"], job["location"], job["age"],
            make_link_formula(link), date.today().strftime("%m/%d/%Y"), job["platform"],
            description, doc_cell, questions_count
        ]

    for attempt in range(3):
        try:
            sheet.append_row(row, value_input_option="USER_ENTERED", table_range="A1")
            seen_links.add(link)
            seen_jobs.add(key)
            STATS["saved"] += 1
            log("SAVE", f"saved {who} -> {short(link)} (description {len(description)} chars)")
            return True
        except Exception as e:
            if attempt == 2:
                STATS["bad"] += 1
                log("ERR", f"{who}: could not write the row to the sheet ({str(e)[:120]})")
            else:
                time.sleep(2)
    return False


async def dismiss_overlays(page):
    """Dismiss common cookie consent dialogs, popups, and GDPR modals if present."""
    dismiss_js = """() => {
        const selectors = [
            'button#onetrust-accept-btn-handler',
            'button#onetrust-reject-all-handler',
            'button.onetrust-close-btn-handler',
            '#onetrust-banner-sdk button',
            'button[id*="cookie" i]',
            'button[class*="cookie" i]',
            'button[aria-label*="cookie" i]',
            'button[aria-label="Close" i]',
            'button[aria-label="close" i]',
            'button:has-text("Accept All")',
            'button:has-text("Accept all")',
            'button:has-text("Accept Cookies")',
            'button:has-text("Accept cookies")',
            'button:has-text("Allow All")',
            'button:has-text("Allow Cookies")',
            'button:has-text("I Agree")',
            'button:has-text("I agree")',
            'button:has-text("Got it")',
            'button:has-text("Dismiss")',
            'button:has-text("Close")',
            '.modal-close',
            '.popup-close'
        ];
        for (const s of selectors) {
            try {
                const btn = document.querySelector(s);
                if (btn && btn.offsetParent !== null) {
                    btn.click();
                    return true;
                }
            } catch (e) {}
        }
        return false;
    }"""
    try:
        await page.evaluate(dismiss_js)
    except Exception:
        pass


async def read_description(page):
    """Description text of the page that is open: dismiss overlays first, then JSON-LD, then main text block."""
    await dismiss_overlays(page)
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
    if "just now" in text or "just posted" in text or "today" in text or text == "now":
        return True
    match = re.search(r"(\d+)\+?\s*(d|days?|h|hrs?|hours?|m|mins?|minutes?)\b", text)
    if not match:
        return False
    unit, value = match.group(2), int(match.group(1))
    if unit.startswith("d"):
        return value <= 1
    return unit.startswith("m") or value <= 24


JOB_CONCURRENCY = 3  # offer pages opened at once


def connect_sheet():
    """Sheet connection plus the duplicate-check lists (runs in a thread while the search page loads)."""
    t0 = time.time()
    try:
        sheet = get_google_sheet()
        if sheet:
            ensure_header(sheet)
        seen_links = load_seen_links(sheet) if sheet else set()
        seen_jobs = load_seen_jobs(sheet) if sheet else set()
        log("INFO", f"phase=sheet connected in {time.time() - t0:.1f}s")
        return sheet, seen_links, seen_jobs
    except Exception as e:
        log("WARN", f"Google Sheets connection failed: {e}")
        return None, set(), set()


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
    log("INFO", f"=== Starting Jobgether Scraper | Keyword: '{keyword}' | Location: '{location_slug}' ===")

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
        sheet_task = asyncio.create_task(asyncio.to_thread(connect_sheet))

        for attempt in range(4):
            try:
                log("INFO", f"phase=search keyword='{keyword}' location='{location_slug}' remote=True")
                await page.goto(target_url, wait_until="domcontentloaded", timeout=30000)
                try:
                    await page.wait_for_selector('a[href^="/offer/"]', timeout=12000)
                except Exception:
                    try:
                        text = (await page.title()) + " " + (await page.inner_text("body"))[:500]
                    except Exception:
                        text = ""
                    if looks_like_captcha(page.url, text):
                        STATS["status"] = "BOTCHECK"
                        log("WARN", "the search page shows a bot-check; not trying to get past it")
                    else:
                        log("WARN", "the search page has no job cards")
                last, stable = -1, 0
                for _ in range(20):
                    n = await page.locator('a[href^="/offer/"]').count()
                    stable = stable + 1 if n == last else 0
                    if stable >= 2:
                        break
                    last = n
                    await page.wait_for_timeout(250)
                break
            except Exception as e:
                if attempt == 3:
                    log("ERR", "network error: could not reach jobgether.com after multiple attempts")
                    STATS["bad"] += 1
                    STATS["status"] = "FAILED"
                    summary()
                    sheet_task.cancel()
                    return
                log("WARN", f"network hiccup reaching jobgether.com, retrying ({attempt + 1}/3)")
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
                        if (lower.includes("just now") || lower.includes("just posted") || lower.includes("new")) {
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
            STATS["cards"] = len(cards)
            log("INFO", f"phase=collect {len(cards)} cards read")

            # Collect all cards posted within the last 24 hours
            fresh_cards = [c for c in cards if is_within_24h(c["posted"])]
            STATS["fresh"] = len(fresh_cards)
            log("INFO", f"{len(fresh_cards)} fresh (24h) of {len(cards)} cards")

            # The sheet connection was started while the search page loaded.
            try:
                sheet, seen_links, seen_jobs = await sheet_task
            except Exception as e:
                log("WARN", f"Google Sheets connection failed: {e}")
                sheet, seen_links, seen_jobs = None, set(), set()
            log("INFO", f"duplicate check loaded: {len(seen_links)} links, {len(seen_jobs)} company/title pairs")

            if sheet:
                try:
                    await asyncio.to_thread(sync_offline_jobs, sheet, seen_links, seen_jobs)
                except Exception as sync_e:
                    log("WARN", f"Could not sync cached offline jobs: {sync_e}")

            sem = asyncio.Semaphore(JOB_CONCURRENCY)
            write_lock = asyncio.Lock()

            async def process(card):
                async with sem:
                    # Already in the sheet (same company + title)? Skip before opening the offer page.
                    key = (card["company"].strip().lower(), card["title"].strip().lower())
                    if key[0] and key[1] and key in seen_jobs:
                        STATS["dup"] += 1
                        log("DUP", f"{jid(card['href'])}  {card['company']} | {card['title']}: company+title already in the sheet (before opening the page)")
                        return
                    job_page = None
                    job_start = time.time()
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

                        who = f"{jid(card['href'])}  {card['company']} | {card['title']}"
                        if apply_link and looks_like_captcha(apply_link, ""):
                            STATS["skip"] += 1
                            log("SKIP", f"{who}: apply link looks like a bot-check page")
                            apply_link = None
                        elif not apply_link:
                            STATS["skip"] += 1
                            log("SKIP", f"{who}: could not find the apply link on the page")
                        else:
                            log("JOB", f"{who}: apply link found, read in {time.time() - job_start:.1f}s")

                        if apply_link:
                            # IN THE SAME VISIT: Navigate to apply_link to extract questions and create Doc
                            doc_url, q_count = "", "None"
                            try:
                                apply_tab = await context.new_page()
                                try:
                                    await apply_tab.goto(apply_link, wait_until="domcontentloaded", timeout=25000)
                                    from question_extractor import extract_questions_and_create_doc_async
                                    doc_url, q_count, _ = await extract_questions_and_create_doc_async(
                                        apply_tab, context, card["company"], card["title"], apply_link
                                    )
                                    if doc_url:
                                        log("INFO", f"[{card['company']} | {card['title']}] Form questions synced: {q_count} -> {doc_url}")
                                finally:
                                    await apply_tab.close()
                            except Exception as q_err:
                                log("WARN", f"In-flight question extraction note: {q_err}")

                            job_data = {
                                "company": card["company"], "title": card["title"], "location": "United States",
                                "age": card["posted"], "link": apply_link, "platform": "Jobgether",
                                "description": description,
                                "questions_doc": doc_url,
                                "questions_count": q_count,
                            }
                            written = False
                            if sheet:
                                async with write_lock:
                                    written = await asyncio.to_thread(append_job, sheet, job_data, seen_links, seen_jobs)
                            else:
                                log("WARN", f"{who}: sheet not connected; only in jobgether.xlsx")
                            if not written:
                                # Save to offline cache if Google Sheet write failed or sheet was disconnected
                                save_offline_job(job_data)
                                results.append({
                                    "Job Title": card["title"], "Company": card["company"],
                                    "Job Age": card["posted"], "Location": "United States",
                                    "Apply Link": apply_link,
                                })
                    except Exception as e:
                        STATS["bad"] += 1
                        log("ERR", f"{jid(card['href'])}  error while processing the job: {str(e)[:150]}")
                    finally:
                        if job_page is not None:
                            try:
                                await job_page.close()
                            except Exception:
                                pass

            log("INFO", f"phase=extract {len(fresh_cards)} jobs, {JOB_CONCURRENCY} at a time")
            await asyncio.gather(*(process(c) for c in fresh_cards), return_exceptions=True)

            log("INFO", f"done: {STATS['saved']} rows written directly to Google Sheets")
            summary()

        finally:
            await context.close()

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as e:
        log("ERR", f"fatal: {str(e)[:200]}")
        STATS["status"] = "FAILED"
        summary()
        raise