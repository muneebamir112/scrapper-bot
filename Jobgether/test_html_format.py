import re

# Test with the exact HTML format from the user
html_snippet = '''class="flex items-center gap-1 text-sm md:text-[14px] font-medium md:font-bold text-[#22A5F1]" href="/remote-jobs/company-sitetracker"target="_blank"class="underline" Sitetracker'''

# The text appears at the very end after all attributes
# Pattern: href="/remote-jobs/company-..." followed by attributes and then the text
match = re.search(r'href="[^"]*company[^"]*"[^>]*\s+([A-Za-z][A-Za-z0-9]*)', html_snippet)
if match:
    print(f"Found company: {match.group(1)}")
else:
    # Try simpler: last word after href pattern
    if '/company' in html_snippet:
        parts = html_snippet.split(' ')
        for part in reversed(parts):
            if part and part[0].isupper() and len(part) > 2 and '/' not in part:
                print(f"Found company from end: {part}")