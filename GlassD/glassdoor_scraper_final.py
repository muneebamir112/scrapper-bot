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
import time
from urllib.parse import urlparse
from datetime import datetime, date
from concurrent.futures import ThreadPoolExecutor
from dotenv import load_dotenv

PLATFORM_NAME = "Glassdoor"
_T0 = time.time()
STATS = {"cards": 0, "dup": 0, "skip": 0, "saved": 0, "bad": 0, "status": "ok"}
SKIPS = {}  # why jobs were skipped: reason key -> count


def log(level, msg):
    """Standard log line: YYYY-MM-DD HH:MM:SS [PLATFORM] [LEVEL] Message"""
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{now_str} [{PLATFORM_NAME}] [{level:<5}] {msg}", flush=True)


def short(url):
    """Host + path only, cut to 70 characters. No query string: links can carry tokens."""
    u = urlparse(str(url))
    return (u.netloc + u.path)[:70]


def skip(who, key, text):
    STATS["skip"] += 1
    SKIPS[key] = SKIPS.get(key, 0) + 1
    log("SKIP", f"{who}: {text}")


def summary():
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    elapsed = int(time.time() - _T0)
    why = " ".join(f"{k}={v}" for k, v in SKIPS.items())
    skip_detail = f" ({why})" if why else ""
    print(f"\n{'='*70}\n{now_str} [{PLATFORM_NAME}] Run Completed (Elapsed: {elapsed}s)\n"
          f"Status: {STATS['status']} | Scanned Cards: {STATS['cards']} | "
          f"Saved: {STATS['saved']} | Duplicates: {STATS['dup']} | Skipped: {STATS['skip']}{skip_detail} | Failed: {STATS['bad']}\n"
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

PROFILE_DIR = os.path.join(CURRENT_DIR, "profiles", "glassdoor_profile")

load_dotenv(os.path.join(LAUNCHER_DIR, ".env"))
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "").strip()
SERVICE_ACCOUNT_FILE = os.path.join(LAUNCHER_DIR, "service_account.json")
EXCEL_HEADER = ["Company", "Job Title", "Location", "Job Age", "Application Link"]

def get_google_sheet():
    """Connect to the configured Google Sheet's first worksheet. Returns None (and
    prints why) if the service account file is missing or the connection fails, so
    the scraper can keep running and just skip the Sheets sync."""
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
    """Application links already logged in the Google Sheet ('Job Link' column).
    Used to skip a job only if this exact link was already added in a past run or
    earlier in the current one — multiple postings from the same company are fine
    as long as each has its own link."""
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


OFFLINE_EXCEL = "glassdoor_offline.xlsx"


def save_offline_job(job):
    """Saves a job locally to glassdoor_offline.xlsx only when Google Sheets is unreachable."""
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
            job.get("age", ""), job.get("link", ""), job.get("platform", "Glassdoor"),
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
            platform = r[h_map.get("platform", 5)] or "Glassdoor"
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
  // Check if an element is inside a cookie, consent, banner, or dialog overlay
  function isCookieOrOverlay(node) {
    if (!node || node === document.body) return false;
    try {
      if (node.closest('#onetrust-consent-sdk, #onetrust-banner-sdk, #CybotCookiebotDialog, [id*="cookie" i], [class*="cookie" i], [id*="consent" i], [class*="consent" i], [id*="banner" i], [class*="banner" i], [role="dialog"], [aria-modal="true"], .modal, .popup')) {
        return true;
      }
    } catch (e) {}
    return false;
  }

  // Quick check if text is predominantly a cookie/privacy consent notice
  function isCookieText(text) {
    if (!text) return false;
    const lower = text.toLowerCase();
    const cookieHits = (lower.match(/cookie/g) || []).length;
    const consentHits = (lower.match(/consent|gdpr|tracking|preferences|strictly necessary/g) || []).length;
    if ((cookieHits >= 2 || (cookieHits >= 1 && consentHits >= 1)) && !lower.includes('requirements') && !lower.includes('responsibilities') && !lower.includes('qualifications')) {
      return true;
    }
    if (lower.startsWith('we use cookies') || lower.startsWith('by clicking "accept') || lower.startsWith('by clicking accept')) {
      return true;
    }
    return false;
  }

  // 1. High-priority target selectors for ATS job descriptions
  const specificSelectors = [
    '[data-automation-id="jobPostingDescription"]',
    '[data-automation-id="job-posting-description"]',
    '[class*="job-description" i]',
    '[id*="job-description" i]',
    '[class*="jobDescription" i]',
    '[id*="jobDescription" i]',
    '[class*="job_description" i]',
    '[id*="job_description" i]',
    '[class*="job-details" i]',
    '[class*="jobdetail" i]',
    '[data-testid="job-description"]',
    '[itemprop="description"]'
  ];

  for (const s of specificSelectors) {
    for (const el of document.querySelectorAll(s)) {
      if (isCookieOrOverlay(el)) continue;
      const t = (el.innerText || '').trim();
      if (t.length >= 200 && t.length < 60000 && !isCookieText(t)) {
        return t;
      }
    }
  }

  // 2. Broader groups, explicitly filtering out cookie/consent elements
  const fallbackGroups = [
    'article',
    'main, [role="main"]',
    '[class*="description" i]:not([class*="cookie" i]):not([class*="banner" i]):not([class*="consent" i]), [id*="description" i]:not([id*="cookie" i]):not([id*="banner" i]):not([id*="consent" i])'
  ];

  for (const g of fallbackGroups) {
    let best = '';
    for (const el of document.querySelectorAll(g)) {
      if (isCookieOrOverlay(el)) continue;
      const t = (el.innerText || '').trim();
      if (isCookieText(t)) continue;
      if (t.length > best.length && t.length < 60000) best = t;
    }
    if (best.length >= 400) return best;
  }

  if (document.body) {
    const bodyText = (document.body.innerText || '').trim();
    if (!isCookieText(bodyText)) return bodyText;
  }
  return '';
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
    re.compile(r"accept cookies", re.IGNORECASE),
    re.compile(r"allow all cookies", re.IGNORECASE),
    re.compile(r"allow cookies", re.IGNORECASE),
    re.compile(r"strictly necessary cookies", re.IGNORECASE),
    re.compile(r"performance cookies", re.IGNORECASE),
    re.compile(r"targeting cookies", re.IGNORECASE),
    re.compile(r"functional cookies", re.IGNORECASE),
    re.compile(r"by clicking [\"']?accept[\"']?", re.IGNORECASE),
    re.compile(r"privacy policy.*cookies", re.IGNORECASE),
    re.compile(r"manage preferences", re.IGNORECASE),
    re.compile(r"consent to the use of cookies", re.IGNORECASE),
    re.compile(r"store and/or access information on a device", re.IGNORECASE),
]

JOB_SIGNALS = [
    re.compile(r"requirements?", re.IGNORECASE),
    re.compile(r"responsibilities", re.IGNORECASE),
    re.compile(r"qualifications?", re.IGNORECASE),
    re.compile(r"experience", re.IGNORECASE),
    re.compile(r"about the role", re.IGNORECASE),
    re.compile(r"about the job", re.IGNORECASE),
    re.compile(r"what you(?:'ll| will) do", re.IGNORECASE),
    re.compile(r"what we(?:'re| are) looking for", re.IGNORECASE),
    re.compile(r"skills?", re.IGNORECASE),
    re.compile(r"salary", re.IGNORECASE),
    re.compile(r"benefits?", re.IGNORECASE),
]


def looks_like_cookie_banner(text):
    """Returns True if the text is predominantly a cookie/GDPR consent notice."""
    t = text.strip()
    if not t:
        return False
    cookie_matches = sum(1 for p in COOKIE_PATTERNS if p.search(t))
    job_matches = sum(1 for p in JOB_SIGNALS if p.search(t))
    if cookie_matches >= 2 and job_matches == 0:
        return True
    if cookie_matches >= 3 and job_matches <= 1:
        return True
    if len(t) < 800 and ("cookie" in t.lower() or "consent" in t.lower()) and job_matches == 0:
        return True
    if t.lower().startswith("we use cookies") and job_matches == 0:
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
        skip(who, "no_link", "no apply link")
        return False
    if link in seen_links or (key[0] and key[1] and key in seen_jobs):
        STATS["dup"] += 1
        log("DUP", f"{who}: {'link' if link in seen_links else 'company+title'} already in the sheet (at write)")
        return False
    problem = row_problem(job)
    if problem:
        skip(who, "incomplete", problem)
        return False
    description = clean_description(job.get("description", ""))
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


def dismiss_overlays(page):
    """Dismiss common cookie consent dialogs, popups, and GDPR modals if present."""
    dismiss_js = r"""() => {
        let dismissed = false;
        const selectors = [
            '#onetrust-accept-btn-handler',
            '#onetrust-pc-btn-handler',
            'button#onetrust-accept-btn-handler',
            '.onetrust-close-btn-handler',
            '#onetrust-banner-sdk button',
            '#CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll',
            '#CybotCookiebotDialogBodyButtonAccept',
            '#CybotCookiebotDialogBodyLevelButtonAccept',
            '#didomi-notice-agree-button',
            '.osano-cm-accept-all',
            '.osano-cm-button--type_save',
            '#truste-consent-button',
            '.trustarc-agree-btn',
            '[data-testid="uc-accept-all-button"]',
            '#axeptio_btn_acceptAll',
            'button.cm-btn-success',
            'button.cm-btn-accept',
            'button.cky-btn-accept',
            '.cc-btn.cc-allow',
            '.cc-btn.cc-dismiss',
            '[data-cookiebanner="accept_button"]',
            '.iubenda-cs-accept-btn',
            'button[id*="cookie-accept" i]',
            'button[id*="accept-cookie" i]',
            'button[class*="cookie-accept" i]',
            'button[class*="accept-cookie" i]',
            'button[id*="accept-all" i]',
            'button[class*="accept-all" i]'
        ];
        for (const s of selectors) {
            try {
                const el = document.querySelector(s);
                if (el && (el.offsetParent !== null || el.getBoundingClientRect().height > 0)) {
                    el.click();
                    dismissed = true;
                    break;
                }
            } catch (e) {}
        }

        if (!dismissed) {
            const acceptRegex = /^(accept\s+all\s+cookies|accept\s+all|accept\s+cookies|allow\s+all\s+cookies|allow\s+all|allow\s+cookies|i\s+accept|i\s+agree|agree\s+&\s+close|agree\s+and\s+continue|accept\s+and\s+proceed|accept\s+&\s+continue|accept\s+recommended|agree|got\s+it|dismiss|ok)$/i;
            const acceptSubRegex = /(accept\s+(all\s+)?cookies|allow\s+(all\s+)?cookies|accept\s+all|allow\s+all|agree\s+to\s+all|accept\s+recommended)/i;
            const candidates = Array.from(document.querySelectorAll('button, a[role="button"], a.btn, [role="button"], input[type="button"], input[type="submit"]'));
            for (const el of candidates) {
                try {
                    const text = (el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || '').trim();
                    if (!text || text.length > 50) continue;
                    if (acceptRegex.test(text) || acceptSubRegex.test(text)) {
                        if (el.offsetParent !== null || el.getBoundingClientRect().height > 0) {
                            el.click();
                            dismissed = true;
                            break;
                        }
                    }
                } catch (e) {}
            }
        }

        // Check shadow roots
        if (!dismissed) {
            try {
                for (const host of document.querySelectorAll('*')) {
                    if (host.shadowRoot) {
                        const btn = host.shadowRoot.querySelector('button[data-testid="uc-accept-all-button"], button#accept, button.accept, [data-testid="uc-accept-all-button"]');
                        if (btn) {
                            btn.click();
                            dismissed = true;
                            break;
                        }
                    }
                }
            } catch (e) {}
        }

        // Force remove persistent overlays/backdrops so they don't cover the real description
        const bannerSelectors = [
            '#onetrust-consent-sdk',
            '#onetrust-banner-sdk',
            '#CybotCookiebotDialog',
            '.osano-cm-window',
            '.didomi-popup-container',
            '[id*="cookie-banner" i]',
            '[class*="cookie-banner" i]',
            '[id*="cookie-notice" i]',
            '[class*="cookie-notice" i]',
            '[id*="cookie-consent" i]',
            '[class*="cookie-consent" i]',
            '[id*="consent-banner" i]',
            '[class*="consent-banner" i]',
            '.cky-consent-container'
        ];
        for (const bSel of bannerSelectors) {
            try {
                document.querySelectorAll(bSel).forEach(bEl => {
                    bEl.remove();
                    dismissed = true;
                });
            } catch (e) {}
        }

        try {
            if (document.body) {
                document.body.style.overflow = 'auto';
                document.body.classList.remove('modal-open', 'no-scroll', 'ot-modal-open');
            }
            if (document.documentElement) {
                document.documentElement.style.overflow = 'auto';
            }
        } catch (e) {}

        return dismissed;
    }"""
    try:
        page.evaluate(dismiss_js)
    except Exception:
        pass

    pw_selectors = [
        "button:has-text('Accept all cookies')",
        "button:has-text('Accept All Cookies')",
        "button:has-text('Accept all')",
        "button:has-text('Accept All')",
        "button:has-text('Accept Cookies')",
        "button:has-text('Allow all cookies')",
        "button:has-text('Allow All')",
        "button:has-text('I Agree')",
        "button:has-text('I agree')",
        "button:has-text('Got it')",
        "button:has-text('Dismiss')",
    ]
    for sel in pw_selectors:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=400):
                loc.click(timeout=1000)
                time.sleep(0.5)
                break
        except Exception:
            pass


def read_description(page):
    """Description text of the page that is open: dismiss overlays first, then JSON-LD, then the main text block."""
    dismiss_overlays(page)
    candidates = []
    try:
        candidates.append(description_from_jsonld(page.evaluate(JSONLD_JS)))
    except Exception:
        pass
    try:
        candidates.append(page.evaluate(BLOCK_JS))
    except Exception:
        pass
    desc = pick_description(candidates)

    if not desc or looks_like_cookie_banner(desc):
        try:
            dismiss_overlays(page)
            time.sleep(1)
            candidates.append(page.evaluate(BLOCK_JS))
            desc = pick_description(candidates)
        except Exception:
            pass

    return desc


def connect_sheet():
    """Sheet connection plus the duplicate-check lists (runs in a thread while the browser works)."""
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
            log("ERR", f"could not read the inputs: {e}")
            STATS["status"] = "FAILED"
            summary()
            return
    
    # Create Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Glassdoor Jobs"
    ws.append(EXCEL_HEADER)

    # Connect to Google Sheets in a background thread while the browser starts and the search runs
    # (jobs are pushed there live, row by row, as they're scraped). The thread is finished before
    # the first row is written, so the sheet is never used by two threads at once.
    sheet_future = ThreadPoolExecutor(max_workers=1).submit(connect_sheet)
    log("INFO", f"=== Starting Glassdoor Scraper | Keyword: '{job_title}' | Location: '{location}' ===")

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
            log("ERR", f"Chrome launch failed: {e}")
            STATS["status"] = "FAILED"
            summary()
            return

        page = context.pages[0] if context.pages else context.new_page()

        try:
            # Step 1: Navigate to Glassdoor Jobs
            log("INFO", "step=open glassdoor.com")
            for attempt in range(4):
                try:
                    page.goto("https://www.glassdoor.com/Job/index.htm", timeout=60000, wait_until="domcontentloaded")
                    break
                except Exception as e:
                    if attempt == 3:
                        log("ERR", "network error: could not reach glassdoor.com after multiple attempts")
                        STATS["status"] = "FAILED"
                        return
                    log("WARN", f"network hiccup reaching glassdoor.com, retrying ({attempt + 1}/3)")
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
                log("WARN", "search inputs not found, trying alternative selectors")
                page.wait_for_selector("input[placeholder*='job title'], input[placeholder*='keyword']", timeout=15000)
                page.wait_for_selector("input[placeholder*='location'], input[name='location']", timeout=15000)
            
            log("INFO", "step=fill the search form")
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
            log("INFO", "step=wait for results")
            page.wait_for_selector("li[data-test='jobListing']", timeout=30000)
            page.wait_for_timeout(3000)
            
            # Step 5: Apply Remote filter and ensure others are unchecked
            log("INFO", "step=apply the Remote filter")
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
                log("WARN", "Remote button not found")

            if "remoteWorkType=1" not in page.url:
                log("WARN", "the Remote filter did not apply (remoteWorkType=1 is missing from the address); stopping so no on-site jobs are saved")
                STATS["status"] = "STOPPED"
                return

            # Step 5b: Apply Glassdoor's own "Date posted: Last day" filter via
            # its fromAge=1 URL param, instead of loading every listing (most
            # of which are >24h old) and filtering them out client-side after
            # the fact - this cuts the number of jobs to load and process from
            # dozens/hundreds down to just the ones already within 24h.
            log("INFO", "step=apply the last-24h filter")
            try:
                sep = "&" if "?" in page.url else "?"
                date_filtered_url = f"{page.url}{sep}fromAge=1"
                page.goto(date_filtered_url, timeout=30000, wait_until="domcontentloaded")
                page.wait_for_timeout(2000)
            except Exception as e:
                log("WARN", f"could not apply the date filter, falling back to client-side filtering only: {e}")

            date_filter_on = "fromAge=1" in page.url
            if not date_filter_on:
                log("WARN", "the 'last day' filter is not in the address; jobs without a readable age will be skipped")

            # Step 6: Load all jobs by clicking "Show more jobs" or scrolling
            log("INFO", "step=load all results")
            max_load_attempts = 50
            load_attempts = 0
            last_count = 0
            consecutive_no_increase = 0
            
            while load_attempts < max_load_attempts and consecutive_no_increase < 3:
                round_start = time.time()
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
                    click_start = time.time()
                    try:
                        btn = page.query_selector(selector)
                        if btn and btn.is_visible():
                            btn.click(timeout=5000)
                            page.wait_for_timeout(3000)
                            clicked = True
                            break
                    except Exception as e:
                        log("WARN", f"'show more' click with {selector[:45]} did not work after {time.time() - click_start:.1f}s: {str(e).splitlines()[0][:80]}")
                        continue

                # Check if job count increased
                new_count = len(page.query_selector_all("li[data-test='jobListing']"))
                if new_count > current_count:
                    consecutive_no_increase = 0
                else:
                    consecutive_no_increase += 1

                load_attempts += 1
                log("INFO", f"load round {load_attempts}: {new_count} listings, {time.time() - round_start:.1f}s, show-more clicked={clicked}")

            job_cards = page.query_selector_all("li[data-test='jobListing']")
            total = len(job_cards)
            STATS["cards"] = total
            log("INFO", f"phase=extract {total} listings found")

            # The sheet connection was started while the browser worked.
            sheet, seen_links, seen_jobs = sheet_future.result()
            log("INFO", f"duplicate check loaded: {len(seen_links)} links, {len(seen_jobs)} company/title pairs")

            if sheet:
                try:
                    sync_offline_jobs(sheet, seen_links, seen_jobs)
                except Exception as sync_e:
                    log("WARN", f"Could not sync cached offline jobs: {sync_e}")

            # Step 8: Process all jobs
            processed = 0
            for idx, card in enumerate(job_cards, 1):
                job_start = time.time()
                who = f"#{idx}"
                try:
                    link_el = card.query_selector("a[data-test='job-link']")
                    card_title = link_el.inner_text().strip().splitlines()[0][:60] if link_el else ""
                    if card_title:
                        who = f"#{idx} {card_title}"
                except Exception:
                    pass
                try:
                    # Skip Easy Apply
                    if card.query_selector("[data-test='easy-apply'], .easy-apply, div:has-text('Easy Apply')"):
                        skip(who, "easy_apply", "Easy Apply (no employer link)")
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
                        skip(who, "older_24h", f"older than 24h ({job_age})")
                        continue

                    # Get job URL
                    link = card.query_selector("a[data-test='job-link'], a[href*='job/']")
                    if not link:
                        skip(who, "no_card_link", "no job link on the card")
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
                    who = f"#{idx} {company} | {job_t}"
                    log("JOB", f"{who}: page read in {time.time() - job_start:.1f}s (description {len(description)} chars)")

                    # Same job already in the sheet from another platform (or
                    # earlier in this run)? Skip before spending time on the
                    # Apply-click flow below.
                    job_key = (company.strip().lower(), job_t.strip().lower())
                    if job_key in seen_jobs:
                        STATS["dup"] += 1
                        log("DUP", f"{who}: company+title already in the sheet (before the Apply click)")
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
                        skip(who, "no_apply_button", "no 'Apply on employer site' button")
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
                        skip(who, "blank_tab", "the application tab never left about:blank")
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
                        skip(who, "captcha", f"bot-check/verification page ({short(app_url)})")
                        job_page.close()
                        continue

                    # "Apply on employer site" sometimes just opens Glassdoor's
                    # own apply-tracking/interstitial page rather than truly
                    # leaving Glassdoor - not a page the applicant can actually
                    # submit an application from, so it doesn't count as
                    # ready-to-apply either.
                    if "glassdoor.com" in app_url:
                        skip(who, "still_glassdoor", f"still on Glassdoor, never reached the employer's page ({short(app_url)})")
                        job_page.close()
                        continue

                    # Duplicate check: skip only if this exact application link was
                    # already added (in this run or a past one). Multiple postings from
                    # the same company are fine as long as each has its own link.
                    if app_url in seen_links:
                        STATS["dup"] += 1
                        log("DUP", f"{who}: application link already in the sheet (before the write)")
                        job_page.close()
                        continue

                    # IN THE SAME VISIT: Extract questions from job_page before closing it!
                    doc_url, q_count = "", "None"
                    try:
                        from question_extractor import extract_questions_and_create_doc_sync
                        doc_url, q_count, _ = extract_questions_and_create_doc_sync(
                            job_page, context, company, job_t, app_url
                        )
                        if doc_url:
                            log("INFO", f"[{company} | {job_t}] Form questions synced: {q_count} -> {doc_url}")
                    except Exception as q_err:
                        log("WARN", f"In-flight question extraction note: {q_err}")

                    saved = False
                    job_dict = {
                        "company": company, "title": job_t, "location": loc, "age": job_age,
                        "link": app_url, "platform": "Glassdoor", "description": description,
                        "questions_doc": doc_url, "questions_count": q_count,
                    }
                    if not sheet:
                        log("WARN", f"{who}: sheet not connected; only in glassdoor.xlsx")
                    else:
                        saved = append_job(sheet, job_dict, seen_links, seen_jobs)
                    if not saved:
                        save_offline_job(job_dict)
                    else:
                        ws.append([company, job_t, loc, job_age, app_url])
                        processed += 1
                    job_page.close()
                    
                    # Test mode limit
                    if TEST_LIMIT and processed >= TEST_LIMIT:
                        log("INFO", f"test limit reached ({TEST_LIMIT} jobs); stopping early")
                        break
                    
                except Exception as e:
                    STATS["bad"] += 1
                    log("ERR", f"{who}: error while processing the job: {str(e)[:150]}")
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
            
            log("INFO", f"done: {STATS['saved']} rows written directly to Google Sheets")
            
        except Exception as e:
            log("ERR", f"fatal: {str(e)[:200]}")
            STATS["status"] = "FAILED"
            import traceback
            traceback.print_exc()
        finally:
            context.close()
            summary()

if __name__ == "__main__":
    main()
