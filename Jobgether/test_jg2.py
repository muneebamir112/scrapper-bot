from patchright.sync_api import sync_playwright

def test_jobgether():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=['--disable-blink-features=AutomationControlled'])
        page = browser.new_page(user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64)')
        page.goto('https://jobgether.com/search-offers?location=united-states', wait_until='domcontentloaded', timeout=60000)
        page.wait_for_timeout(5000)
        
        cards = page.query_selector_all('a[href^="/offer/"]')
        if cards:
            parent = cards[0].evaluate_handle('el => el.closest("a").parentElement.parentElement.parentElement')
            print(parent.evaluate('el => el.outerHTML'))
        browser.close()

if __name__ == '__main__':
    test_jobgether()
