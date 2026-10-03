import asyncio
from src.scraper.browser import BrowserSession
from src.scraper.bookings import fetch_bookings_list

async def main():
    async with BrowserSession(headless=True) as page:
        await fetch_bookings_list(page, tab="Cancelled", max_rows=5)
        
        # Dump all elements with class containing 'page', 'pager', 'select', 'dropdown'
        elements = await page.evaluate("""() => {
            const elList = [];
            // Look for any elements with text 'of 137' or '137 items' or 'items' or '1000'
            const allElements = document.querySelectorAll('*');
            for (let el of allElements) {
                if (el.children.length === 0 && (el.innerText || '').includes('items')) {
                    elList.push({
                        tag: el.tagName,
                        text: el.innerText,
                        parentTag: el.parentElement ? el.parentElement.tagName : null,
                        parentClass: el.parentElement ? el.parentElement.className : null,
                        parentHtml: el.parentElement ? el.parentElement.outerHTML.substring(0, 300) : null
                    });
                }
            }
            return elList;
        }""")
        
        import json
        print("Elements containing 'items':")
        print(json.dumps(elements, indent=2))

asyncio.run(main())
