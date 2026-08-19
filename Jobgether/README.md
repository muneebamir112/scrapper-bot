# Jobgether Scraper

A Python script that scrapes job listings from [jobgether.com](https://jobgether.com/home) using Playwright with Chrome remote debugging.

## Prerequisites

- Python 3.13+
- Google Chrome installed at `C:\Program Files\Google\Chrome\Application\chrome.exe`
- Chrome running with remote debugging enabled

## Setup Chrome with Remote Debugging

Start Chrome with the following command:

```cmd


```

Login to jobgether.com in the browser before running the script.

## Installation

Install dependencies:

```cmd
pip install playwright pandas openpyxl
```

## Usage

Run the script:

```cmd
python jobgether_scraper.py
```

The script will prompt you for:

1. **Job keyword** - Enter the job title/role to search for (e.g., "full stack engineer")
2. **Number of scrolls** - Enter how many times to scroll down to load more jobs (press Enter for default 15)

## Output

The script saves the results to a timestamped Excel file (e.g., `jobs_20260517_013439.xlsx`) with the following columns:

- **Job Title** - The job position title
- **Company** - The company name
- **Job Age** - How long ago the job was posted
- **Location** - Job location (Remote, Canada, etc.)
- **Apply Link** - Direct link to apply for the job

## Features

- Opens jobgether.com in a new tab
- Applies keyword filter through the UI
- Scrolls to load the specified number of times
- Filters out jobs older than 15 days
- Opens each job in a new tab to extract detailed information
- Handles job cards that may have age information in text

## Troubleshooting

- Ensure Chrome is running with remote debugging on port 9222
- Make sure you're logged into jobgether.com before running the script
- If the script fails to find elements, the website layout may have changed

## Files

- `jobgether_scraper.py` - Main scraper script