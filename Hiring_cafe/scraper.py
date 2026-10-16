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
import re
import json
import time
from urllib.parse import urlparse, quote
from patchright.async_api import async_playwright
import openpyxl
import gspread
from google.oauth2.service_account import Credentials
from datetime import date

TAG = "HC"
_T0 = time.time()
STATS = {"cards": 0, "pages": 0, "fresh": 0, "dup": 0, "skip": 0, "saved": 0, "bad": 0}


def log(level, msg):
    """One log line: [platform] +seconds-since-start LEVEL message (the launcher adds the clock time)."""
    print(f"[{TAG}] +{time.time() - _T0:5.1f}s {level:<5} {msg}", flush=True)


def short(url):
    """Host + path only, cut to 70 characters. No query string: links can carry tokens."""
    u = urlparse(str(url))
    return (u.netloc + u.path)[:70]


def jid(link):
    """Short id of a job link, to follow one job through the log."""
    return urlparse(str(link)).path.rstrip("/")[-8:]


def summary(status="ok"):
    log("SUM", f"status={status} cards={STATS['cards']} pages={STATS['pages']} fresh={STATS['fresh']} "
               f"dup={STATS['dup']} skip={STATS['skip']} saved={STATS['saved']} bad={STATS['bad']} "
               f"elapsed={time.time() - _T0:.0f}s")

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
from dotenv import load_dotenv
load_dotenv(os.path.join(LAUNCHER_DIR, ".env"))
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1Kva2y5-54LXBWMTzNk_xp524ZE7N-CWiqL3VGATUZVM")
SERVICE_ACCOUNT_FILE = os.path.join(LAUNCHER_DIR, "service_account.json")


def get_google_sheet():
    """Connect to Sheet1 of the shared Google Sheet. Returns None (and prints why)
    if the service account file is missing or the connection fails, so the
    scraper can keep running and just skip the Sheets sync."""
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
            log("WARN", f"could not format headers: {e}")


def load_seen_links(worksheet):
    """Application links already logged in the sheet (column F, 'Job Link'),
    used to skip a job only if this exact link was already added before."""
    try:
        link_col = worksheet.col_values(6)
    except Exception as e:
        log("WARN", f"could not read existing job links from the sheet: {e}")
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
        log("WARN", f"could not read existing company/title pairs from the sheet: {e}")
        return set()
    seen = set()
    for company, title in zip(company_col[1:], title_col[1:]):
        company = company.strip().lower()
        title = title.strip().lower()
        if company and title:
            seen.add((company, title))
    return seen


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


async def _wait_for_links_to_settle(page, max_ms=3000, step_ms=500):
    """Wait until the number of job links on the page stops changing, so a partly rendered page is never read."""
    last = -1
    for _ in range(max_ms // step_ms):
        count = await page.locator('a[href^="/job/"]').count()
        if count == last:
            return
        last = count
        await page.wait_for_timeout(step_ms)


# Clicks the first "Apply" button (not "Applied"); returns whether one was found.
APPLY_CLICK_JS = """() => {
    for (const btn of document.querySelectorAll('button')) {
        if (btn.textContent && /Apply(?!ed)/.test(btn.textContent)) { btn.click(); return true; }
    }
    return false;
}"""


PAGE_CONCURRENCY = 3   # list pages read at once
JOB_CONCURRENCY = 3    # job pages opened at once

# Reads every job card on the open list page: link plus the age text shown on the card.
CARDS_JS = '''els => els.map(e => {
    let card = e.closest("div");
    for (let i=0; i<3; i++) { if (card && card.parentElement) card = card.parentElement; }
    return {href: e.href, text: card ? card.innerText : ""};
})'''


async def read_list_page(context, url, shared_page=None):
    """Read one list page. Returns (state, raw_card_count, [(link, age)], reason, timing text);
    state is ok / empty / error. Opens its own tab unless shared_page is given (page 1,
    already loaded). Never touches shared state, so pages can be read at the same time."""
    t0 = time.time()
    marks = {}

    def took():
        # "3.4s (goto 1.2 / first card 2.0 / settled 3.1)": where the time of this page went
        parts = " / ".join(f"{k} {v:.1f}" for k, v in marks.items())
        return f"{time.time() - t0:.1f}s" + (f" ({parts})" if parts else "")

    pg = shared_page or await context.new_page()
    try:
        if shared_page is None:
            await pg.goto(url, wait_until="domcontentloaded", timeout=60000)
        marks["goto"] = time.time() - t0
        try:
            await pg.wait_for_selector('a[href^="/job/"]', timeout=10000)
        except Exception:
            try:
                text = (await pg.title()) + " " + (await pg.inner_text("body"))[:500]
            except Exception:
                text = ""
            if looks_like_captcha(pg.url, text):
                return "error", 0, [], "bot-check page", took()
            return "empty", 0, [], "", took()
        marks["cards"] = time.time() - t0
        await _wait_for_links_to_settle(pg)
        marks["settle"] = time.time() - t0
        data = await pg.eval_on_selector_all('a[target="_blank"][rel="noopener noreferrer"][href^="/job/"]', CARDS_JS)
        links = []
        seen = set()
        for item in data:
            href = item['href']
            if href in seen:
                continue
            seen.add(href)
            text = item['text'].lower()
            age_text = None
            if "just now" in text:
                age_text = "Just now"
            elif "today" in text:
                age_text = "Today"
            else:
                match = re.search(r'(?:posted)?\s*(\d+)\s*(mo|months?|w|weeks?|d|days?|h|hrs?|hours?|m|mins?|minutes?)(?:\s*ago)?\b', text)
                if match:
                    found = re.sub(r"^\s*posted\s*", "", match.group(0)).strip()
                    if is_within_24h(found):
                        age_text = found if found.endswith("ago") else found + " ago"
            if age_text:
                links.append((href, age_text))
        return ("ok" if data else "empty"), len(data), links, "", took()
    except Exception as e:
        return "error", 0, [], str(e)[:120], took()
    finally:
        if shared_page is None:
            try:
                await pg.close()
            except Exception:
                pass


def connect_sheet():
    """Sheet connection plus the duplicate-check lists (runs in a thread while the list pages load)."""
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


def make_link_formula(url):
    """Wrap a URL in a HYPERLINK() formula so it renders as a clickable link
    in the sheet (gspread's default FORMATTED_VALUE reads still return the
    plain URL text, so nothing downstream that reads the sheet is affected)."""
    escaped = url.replace('"', '""')
    return f'=HYPERLINK("{escaped}", "{escaped}")'


import html as _html


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
    description = clean_description(job.get("description", ""))
    row = [
        "=ROW()-1", job["company"], job["title"], job["location"], job["age"],
        make_link_formula(link), date.today().strftime("%m/%d/%Y"), job["platform"],
        description,
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
        sheet_task = asyncio.create_task(asyncio.to_thread(connect_sheet))

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
                log("INFO", "clearing stored site settings")
                await page.goto("https://hiringcafe.com/", wait_until="domcontentloaded", timeout=60000)
                await page.evaluate("window.localStorage.clear(); window.sessionStorage.clear();")

                log("INFO", f"phase=search keyword='{keyword}' location='{location}' remote=True")
                await page.goto(search_url, wait_until="domcontentloaded", timeout=60000)
                try:
                    await page.wait_for_selector('a[href^="/job/"]', timeout=10000)
                except Exception:
                    pass

                # If a filter chip shows the wrong setting (Hybrid / Onsite / Field), remove it
                removed_chip = await page.evaluate('''() => {
                    let removed = false;
                    const chips = Array.from(document.querySelectorAll('div.bg-gray-200.rounded-md'));
                    chips.forEach(chip => {
                        const txt = chip.textContent || '';
                        if (txt.includes('Hybrid') || txt.includes('Onsite') || txt.includes('Field') || (!txt.includes('United States') && !txt.includes('Remote'))) {
                            const trash = chip.querySelector('svg.text-red-500');
                            if (trash) {
                                const btn = trash.closest('button');
                                if (btn) { btn.click(); removed = true; }
                            }
                        }
                    });
                    return removed;
                }''')
                if removed_chip:
                    await page.wait_for_timeout(2000)
                log("INFO", "search page ready")
                break
            except Exception as e:
                if attempt == 3:
                    log("ERR", "network error: could not reach hiringcafe.com after multiple attempts")
                    STATS["bad"] += 1
                    summary("FAILED")
                    sheet_task.cancel()
                    return
                log("WARN", f"network hiccup reaching hiringcafe.com, retrying ({attempt + 1}/3)")
                await asyncio.sleep(2)

        # Hiring Cafe uses URL pagination. We'll collect up to 50 links. Page 1 is already open;
        # the next pages are read PAGE_CONCURRENCY at a time but used strictly in page order.
        all_job_links = []
        link_ages = {}

        def add_page(n, raw, items, secs):
            for href, age in items:
                link_ages.setdefault(href, age)
            all_job_links.extend(h for h, _ in items)
            all_job_links[:] = dict.fromkeys(all_job_links)
            STATS["cards"] += raw
            STATS["pages"] += 1
            log("PAGE", f"{n}: {len(items)} fresh of {raw} cards, total {len(all_job_links)} [{secs}]")
            return len(all_job_links) >= 50

        log("INFO", "phase=collect")
        state, raw, items, why, secs = await read_list_page(context, None, page)
        if state == "error":
            log("WARN", f"page 1 could not be read ({why}); stopping pagination")
        elif state == "empty":
            log("INFO", "no jobs on page 1")
        elif add_page(1, raw, items, secs):
            log("INFO", f"collected enough links ({len(all_job_links)}); stopping pagination")
        else:
            current_url = page.url
            base_url = current_url.split('&page=')[0] if 'page=' in current_url else current_url
            tasks = {}

            async def fetch(n):
                url = f"{base_url}&page={n}"
                res = await read_list_page(context, url)
                if res[0] == "error":
                    res = await read_list_page(context, url)
                return res

            n = 2
            bad_in_a_row = 0
            while True:
                for k in range(n, n + PAGE_CONCURRENCY):
                    if k not in tasks:
                        tasks[k] = asyncio.create_task(fetch(k))
                state, raw, items, why, secs = await tasks[n]
                if state == "empty":
                    later = await asyncio.gather(*(tasks[k] for k in range(n + 1, n + PAGE_CONCURRENCY) if k in tasks))
                    if any(r[0] == "ok" for r in later):
                        log("WARN", f"page {n} came back empty but later pages have jobs; reading it again")
                        state, raw, items, why, secs = await read_list_page(context, f"{base_url}&page={n}")
                    else:
                        log("INFO", f"no more jobs on page {n}; end of the list")
                        break
                if state == "ok":
                    bad_in_a_row = 0
                    if add_page(n, raw, items, secs):
                        log("INFO", f"collected enough links ({len(all_job_links)}); stopping pagination")
                        break
                else:
                    bad_in_a_row += 1
                    log("WARN", f"page {n} was skipped ({state}{': ' + why if why else ''})")
                    if bad_in_a_row >= 3:
                        log("WARN", "three pages in a row could not be read; stopping pagination")
                        break
                n += 1

            for t in tasks.values():
                t.cancel()
            await asyncio.gather(*tasks.values(), return_exceptions=True)

        try:
            await page.goto("about:blank")
        except Exception:
            pass

        # Remove duplicates across pages
        all_job_links = list(dict.fromkeys(all_job_links))
        total_jobs = len(all_job_links)
        STATS["fresh"] = total_jobs
        log("INFO", f"phase=collect done: {total_jobs} fresh links from {STATS['pages']} pages")

        for link in all_job_links:
            log("JOB", f"{jid(link)}  [{link_ages.get(link, '')}] {short(link)}")

        if total_jobs == 0:
            log("INFO", "no fresh jobs found; exiting")
            summary()
            sheet_task.cancel()
            return

        # Ask how many to extract only when run from a real terminal (input() would block
        # forever when the launcher starts this script without one). Default: all collected.
        try:
            interactive = sys.stdin.isatty()
        except Exception:
            interactive = False

        if interactive:
            try:
                num_to_extract = min(int(input(f"How many jobs to extract? (1-{total_jobs}): ")), total_jobs)
            except Exception:
                num_to_extract = total_jobs
        else:
            num_to_extract = total_jobs

        # The sheet connection was started while the list pages loaded.
        try:
            sheet, seen_links, seen_jobs = await sheet_task
        except Exception as e:
            log("WARN", f"Google Sheets connection failed: {e}")
            sheet, seen_links, seen_jobs = None, set(), set()
        log("INFO", f"duplicate check loaded: {len(seen_links)} links, {len(seen_jobs)} company/title pairs")

        log("INFO", f"phase=extract {num_to_extract} jobs, {JOB_CONCURRENCY} at a time")
        sem = asyncio.Semaphore(JOB_CONCURRENCY)
        write_lock = asyncio.Lock()

        async def process(job_url):
            async with sem:
                try:
                    job_data = await extract_job_data(context, job_url, link_ages.get(job_url, ""), seen_jobs)
                except Exception as e:
                    STATS["bad"] += 1
                    log("ERR", f"{jid(job_url)}  extraction error: {str(e)[:150]}")
                    return
            if job_data is None:
                return
            jobs.append(job_data)
            if sheet:
                # one write at a time: the duplicate check, the append and the seen_* update stay together
                async with write_lock:
                    await asyncio.to_thread(append_job, sheet, {
                        "company": job_data["company"], "title": job_data["title"], "location": job_data["location"],
                        "age": job_data["posted"], "link": job_data["apply_link"], "platform": "Hiring Cafe",
                        "description": job_data.get("description", ""),
                    }, seen_links, seen_jobs)
            else:
                log("WARN", f"{jid(job_url)}  sheet not connected; {job_data['company']} | {job_data['title']} is only in jobs.xlsx")

        await asyncio.gather(*(process(u) for u in all_job_links[:num_to_extract]), return_exceptions=True)

        await save_to_xlsx(jobs)
        try:
            await context.close()
        except Exception:
            pass
        log("INFO", f"done: {STATS['saved']} rows written to the sheet, {len(jobs)} jobs in jobs.xlsx")
        summary()

async def extract_job_data(context, job_url, age_text, seen_jobs):
    new_page = await context.new_page()
    popups = []
    new_page.on("popup", lambda p: popups.append(p))  # Playwright needs a Python function, not a bound builtin
    try:
        await new_page.goto(job_url, wait_until="domcontentloaded", timeout=60000)

        # The page is ready once the Apply button is in the page (title and description render with it).
        # It only has to exist: the site keeps hidden duplicate buttons (mobile layout) that sort first.
        apply_visible = True
        try:
            await new_page.locator("button").filter(has_text=re.compile(r"Apply(?!ed)")).first.wait_for(state="attached", timeout=10000)
        except Exception:
            apply_visible = False

        # hiring.cafe's page <title> is consistently formatted as
        # "<Job Title> at <Company> — <Location>", which is far less prone
        # to breaking on a redesign than matching specific CSS classes.
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

        posted = age_text  # from the job card, where it was already checked as fresh

        # Already in the sheet (same company + title)? Skip before any Apply click.
        if company and title and (company.lower(), title.lower()) in seen_jobs:
            STATS["dup"] += 1
            log("DUP", f"{jid(job_url)}  {company} | {title}: company+title already in the sheet (before opening Apply)")
            await new_page.close()
            return None

        board_description = await read_description(new_page)
        employer_description = ""

        # A job is only worth saving if we can confirm a real, external,
        # ready-to-apply link - the hiring.cafe listing page itself isn't
        # one, so this skips the job when a genuine apply link can't be captured.
        def _is_ready_to_apply(link, text):
            if not link or link == "about:blank":
                return False
            if "hiring.cafe" in link:
                return False
            if looks_like_captcha(link, text):
                return False
            return True

        apply_link = None
        reason = "" if apply_visible else "Apply button not found on the page"
        try:
            url_before_click = new_page.url
            popup = None
            clicked = False
            for attempt in range(2):
                popup = None
                clicked = False
                if apply_visible:
                    try:
                        # expect_popup hands back exactly the tab THIS click opened
                        async with new_page.expect_popup(timeout=15000) as pop:
                            clicked = await new_page.evaluate(APPLY_CLICK_JS)
                        popup = await pop.value
                    except Exception:
                        popup = None
                if attempt == 1 or not clicked:
                    break
                if popup is not None:
                    await _wait_for_real_navigation(popup, 15000)
                    if popup.url != "about:blank":
                        break
                    # blank Apply tab: close it and click once more
                    try:
                        await popup.close()
                    except Exception:
                        pass
                else:
                    try:
                        await new_page.wait_for_url(lambda u: u != url_before_click, timeout=3000)
                    except Exception:
                        pass
                    if new_page.url != url_before_click:
                        break

            if popup is not None:
                await _wait_for_real_navigation(popup, 15000)
                candidate_link = popup.url
                try:
                    candidate_text = (await popup.title()) + " " + (await popup.inner_text("body"))[:500]
                except Exception:
                    candidate_text = ""
                employer_description = await read_description(popup)
                try:
                    await popup.close()
                except Exception:
                    pass
                if _is_ready_to_apply(candidate_link, candidate_text):
                    apply_link = candidate_link
                else:
                    reason = f"the Apply tab opened on a rejected link ({short(candidate_link)})"
            elif clicked:
                # No new tab: some sites navigate the same page instead of opening a popup.
                try:
                    await new_page.wait_for_url(lambda u: u != url_before_click, timeout=3000)
                except Exception:
                    pass
                candidate_link = new_page.url
                if candidate_link != url_before_click:
                    try:
                        candidate_text = (await new_page.title()) + " " + (await new_page.inner_text("body"))[:500]
                    except Exception:
                        candidate_text = ""
                    employer_description = await read_description(new_page)
                    if _is_ready_to_apply(candidate_link, candidate_text):
                        apply_link = candidate_link
                    else:
                        reason = f"the page navigated to a rejected link ({short(candidate_link)})"
                else:
                    reason = "the Apply click opened no new tab and the page did not navigate"
            elif apply_visible:
                reason = "the Apply button could not be clicked"
        except Exception as e:
            reason = f"error while capturing the apply link ({e})"

        for extra in popups:
            try:
                await extra.close()
            except Exception:
                pass
        await new_page.close()

        if not apply_link:
            STATS["skip"] += 1
            log("SKIP", f"{jid(job_url)}  {company} | {title}: {reason or 'no apply link captured'}")
            return None

        description = employer_description if len(employer_description) >= 400 else max(employer_description, board_description, key=len)
        return {"title": title.strip(), "company": company.strip(), "location": job_location.strip(),
                "posted": posted.strip(), "job_url": job_url, "apply_link": apply_link,
                "description": description}
    except Exception as e:
        STATS["bad"] += 1
        log("ERR", f"{jid(job_url)}  extraction error: {str(e)[:150]}")
        try:
            await new_page.close()
        except Exception:
            pass
        return None

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
    try:
        asyncio.run(main())
    except Exception as e:
        log("ERR", f"fatal: {str(e)[:200]}")
        summary("FAILED")
        raise