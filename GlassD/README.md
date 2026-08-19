# GlassD - Glassdoor Job Scraper

A Python-based web scraper that extracts job listings from Glassdoor and saves them to an Excel spreadsheet. The scraper launches a stealth-patched Chrome browser (via [patchright](https://github.com/Kaliiiiiiiiii-Vinyzu/patchright-python)) with a persistent profile to handle login, maintain session state, and avoid Glassdoor's bot/captcha checks.

## Features

- **Automated Job Search**: Searches Glassdoor by job title and location
- **Remote Filter**: Automatically applies "Remote only" filter
- **Pagination Handling**: Clicks "See more jobs" until all results are loaded
- **Easy Apply Exclusion**: Skips jobs marked as "Easy Apply"
- **Application URL Extraction**: Visits each job individually, navigates through the application process to capture the final employer application link
- **Excel Export**: Saves all scraped data (Company, Job Title, Location, Application Link) to `glassdoor.xlsx`
- **Session Preservation**: Persists your Glassdoor login in a local Chrome profile
- **Stealth Automation**: Uses `patchright` to avoid the bot/captcha checks that flag plain Playwright/Selenium sessions

## Requirements

### Software Dependencies

- **Python 3.8+**
- **Google Chrome** browser
- **Microsoft Excel** or compatible spreadsheet viewer (for opening `.xlsx` files)

### Python Packages

Install required packages using pip:

```bash
pip install patchright openpyxl gspread google-auth
```

`patchright` is a stealth-patched drop-in replacement for Playwright that removes the CDP artifacts (e.g. the `Runtime.enable` leak) sites like Glassdoor use to fingerprint and captcha-gate automated browsers. It reuses your system-installed Google Chrome (via `channel="chrome"`), so no separate browser binary download is needed.

## Google Sheets Sync (Optional)

Every scraped job is pushed live (row by row) to this Google Sheet, in addition to `glassdoor.xlsx`:
https://docs.google.com/spreadsheets/d/1FsPR9t-BB1GZ6kWfANnrDfq4p9XobDuA1tVB4D2sJDg/edit?gid=0

This uses a Google service account, so the scraper can write to the sheet with no browser login or manual OAuth step at runtime. One-time setup:

1. **Create a Google Cloud project** (or reuse one) at [console.cloud.google.com](https://console.cloud.google.com/).
2. **Enable the Google Sheets API**: APIs & Services → Library → search "Google Sheets API" → Enable.
3. **Create a service account**: APIs & Services → Credentials → Create Credentials → Service Account. Give it any name (e.g. `glassdoor-scraper`).
4. **Create a JSON key**: open the new service account → Keys → Add Key → Create new key → JSON. This downloads a `.json` file.
5. **Rename and place the key** as `service_account.json` in this `GlassD/` folder (same directory as `glassdoor_scraper_final.py`). This file is a secret — it's already listed in `.gitignore` and should never be committed or shared.
6. **Share the Google Sheet** with the service account: open the JSON file, copy the `client_email` value (looks like `xxxx@xxxx.iam.gserviceaccount.com`), then in the Google Sheet click **Share** and add that email as an **Editor**.

Once `service_account.json` is in place and the sheet is shared, just run the scraper normally — it detects the file automatically and syncs jobs as they're scraped. If the file is missing, the scraper prints a notice and continues writing to `glassdoor.xlsx` only (no Sheets sync, no error).

### Duplicate-Link Skipping

Multiple postings from the same company are all kept (e.g. a "DevOps Engineer" and a "Software Developer" role at the same company are both added) — each is a separate application worth its own tailored resume. What's *not* kept is the same job link twice.

- At startup the scraper reads every application link already in the Google Sheet's "Job Link" column (column F, across all past runs) into a "seen" set.
- The final employer application link is only known after a job's Apply flow is followed through, so this check happens right after that link is captured — not before opening the job's tab (there's no reliable way to know the final destination link ahead of time).
- A link is only marked "seen" once its job is actually written to the sheet.
- Skipped duplicates are never written to `glassdoor.xlsx` either.

## Project Structure

```
GlassD/
├── glassdoor_scraper.py        # Original scraper implementation
├── glassdoor_scraper_improved.py      # Improved version with enhanced error handling
├── glassdoor_scraper_final.py         # Final/stable version
├── debug_scraper.py            # Debug version for testing connections
├── inputs.txt                  # Sample input values (job title, location)
├── chrome_profile/             # Persistent Chrome profile (created on first run; holds login session)
├── glassdoor_home.html         # Saved Glassdoor page (reference)
└── README.md                   # This file
```

## Usage

### Step 1: Set Your Search Inputs

Edit `inputs.txt` with two lines:

```
Full Stack Developer
United States
```

1. **Job Title**: e.g., `Full Stack Developer`, `Software Engineer`
2. **Location**: e.g., `United States`, `New York`, `Remote`

### Step 2: Run the Scraper

```bash
python glassdoor_scraper_final.py
```

The script launches a real Chrome window itself (stealth-patched, via `patchright`) using a persistent profile stored in `GlassD/chrome_profile/`. On the **first run**, log into Glassdoor manually in that window before the script proceeds through the search — your session is saved in that profile folder and reused automatically on subsequent runs, so you shouldn't need to log in again.

### Step 3: Let It Run

The script will:
1. Launch the stealth Chrome browser
2. Navigate to Glassdoor
3. Search for the specified job
4. Apply the "Remote only" filter
5. Load all available job listings
6. Visit each job individually
7. Extract company, title, location, and application URL
8. Save results to `glassdoor.xlsx`

### Step 4: View Results

Open the generated `glassdoor.xlsx` file in Excel, Google Sheets, or any compatible spreadsheet application.

## How It Works

### Architecture

1. **Stealth Launch**: Uses `patchright`'s `launch_persistent_context` with `channel="chrome"` to drive your real, installed Chrome — this avoids the CDP fingerprint leaks (e.g. `Runtime.enable`) that anti-bot systems use to detect Playwright/Selenium
2. **Page Navigation**: Directly navigates to Glassdoor's Jobs section
3. **Search**: Fills and submits the job search form
4. **Filtering**: Applies "Remote only" filter through the UI
5. **Pagination**: Repeatedly clicks "See more jobs" until no more results load
6. **Job Processing**:
   - Skips "Easy Apply" jobs
   - Opens each job in a new tab
   - Extracts job metadata (company, title, location)
   - Navigates to the employer application page
   - Captures the final application URL
7. **Data Export**: Writes all collected data to an Excel workbook

### Selector Strategy

The scraper uses multiple CSS selector fallbacks for each UI element to handle Glassdoor's dynamic page structure and A/B testing variations.

## Troubleshooting

### Launch Failed

**Error**: `Launch failed: ...`

**Solutions:**
- Confirm Google Chrome is installed (patchright uses `channel="chrome"`, your system install, not a bundled Chromium)
- Close any Chrome window already using the `chrome_profile` directory — a profile can only be opened by one running Chrome at a time
- Delete/rename `GlassD/chrome_profile/` to reset the profile if it becomes corrupted (you'll need to log into Glassdoor again)

### Element Not Found

If selectors fail to match elements:

**Solutions:**
- Glassdoor may have updated their page structure
- Use `debug_scraper.py` to test connection and inputs
- Open browser DevTools to inspect current element selectors
- Update selector arrays in the scraper with new CSS selectors

### Jobs Not Loading / Timeouts

**Solutions:**
- Increase wait timeouts in the script (search for `timeout=` values)
- Add additional `page.wait_for_timeout()` delays
- Check network connection

### "Apply on employer site" Button Not Found

This happens for jobs that:
- Are Easy Apply only (already filtered out)
- Have a different button text/label
- Already redirect to employer site in the initial job URL

The scraper logs these cases and continues with other jobs.

### Excel File Locked

**Solution:** Close any open instance of `glassdoor.xlsx` before running the scraper.

### Chrome Profile Issues

If Glassdoor doesn't maintain login state, make sure `GlassD/chrome_profile/` isn't being deleted between runs (e.g. by a cleanup script or `.gitignore`-driven tooling) — this folder is what stores the logged-in session.

## Version History

### `glassdoor_scraper.py` (Original)
- Basic CDP connection and search flow
- Simple error handling

### `glassdoor_scraper_improved.py`
- Better page usability validation
- Direct navigation to Jobs page
- More detailed debug logging
- Case-insensitive selector matching

### `glassdoor_scraper_final.py` (Recommended)
- Clean, readable code structure
- Robust selector fallbacks
- Proper resource cleanup
- Better remote filter application

### `debug_scraper.py`
- Minimal scraper for testing CDP connection and input handling
- Stops before any browser automation

## Limitations

- **Rate Limiting**: Rapid requests may trigger Glassdoor's anti-bot mechanisms
- **UI Changes**: Glassdoor frequently updates their page structure; selectors may break
- **First-Run Login**: Requires a one-time manual login in the launched browser window before the saved profile takes over
- **Single Query**: Runs one search at a time (batch mode not implemented)
- **Job Card Limits**: May not capture all jobs if "See more jobs" button disappears prematurely

## Future Improvements

- [ ] Add command-line argument parsing (no interactive prompts)
- [ ] Implement retry logic for failed job processing
- [ ] Add proxy support for rate limit avoidance
- [ ] Export to CSV, JSON in addition to Excel
- [ ] Multi-threaded job processing for faster scraping
- [ ] Config file support for repeated searches
- [ ] Headless mode option
- [ ] Progress bar/status indicators
- [ ] Logging to file

## Disclaimer

This scraper is for **educational purposes and personal use only**. Glassdoor's Terms of Service prohibit automated scraping. Use responsibly and consider:

- Checking Glassdoor's `robots.txt` and Terms of Service
- Adding appropriate delays between requests
- Not overwhelming their servers
- Using official APIs if available
- Respecting rate limits and copyright

The authors assume no responsibility for misuse or violations of Glassdoor's terms.

## License

This project is provided as-is for educational purposes.
