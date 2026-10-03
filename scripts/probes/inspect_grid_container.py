import asyncio
from src.scraper.browser import BrowserSession
from src.scraper.bookings import fetch_bookings_list

async def main():
    async with BrowserSession(headless=True) as page:
        await fetch_bookings_list(page, tab="Cancelled", max_rows=5)
        
        res = await page.evaluate("""() => {
            const container = document.querySelector('.chargingStationBookingDetailsGridContainer');
            if (!container) return { error: 'container not found' };
            
            // Find all tables and pagers inside container
            const tables = container.querySelectorAll('table');
            const pagers = container.querySelectorAll('.k-pager-wrap, .k-grid-pager, div[class*=\"pager\"]');
            const selects = container.querySelectorAll('select');
            const dropdowns = container.querySelectorAll('.k-dropdown, span[role=\"listbox\"]');
            
            return {
                tablesCount: tables.length,
                pagersCount: pagers.length,
                selectsCount: selects.length,
                dropdownsCount: dropdowns.length,
                pagersHTML: Array.from(pagers).map(p => p.outerHTML),
                dropdownsHTML: Array.from(dropdowns).map(d => ({ text: d.innerText, html: d.outerHTML })),
                selectsHTML: Array.from(selects).map(s => ({
                    options: Array.from(s.options).map(o => o.text),
                    value: s.value,
                    html: s.outerHTML
                }))
            };
        }""")
        import json
        print(json.dumps(res, indent=2))

asyncio.run(main())
