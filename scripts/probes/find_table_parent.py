import asyncio
from src.scraper.browser import BrowserSession
from src.scraper.bookings import fetch_bookings_list

async def main():
    async with BrowserSession(headless=True) as page:
        await fetch_bookings_list(page, tab="Cancelled", max_rows=5)
        
        info = await page.evaluate("""() => {
            const table = document.querySelector("table.k-selectable, table:has(tbody tr td)");
            if (!table) return "No table found";
            
            // Check ancestors
            let p = table;
            const ancestors = [];
            while (p && p !== document.body) {
                ancestors.push({
                    tag: p.tagName,
                    id: p.id,
                    className: p.className
                });
                p = p.parentElement;
            }
            
            // Look for any sibling or child that is a pager
            const grid = table.closest('.k-grid') || table.parentElement;
            const pager = grid ? grid.querySelector('.k-pager-wrap, .k-grid-pager, [class*=\"pager\"]') : null;
            
            return {
                ancestors: ancestors,
                gridHTML: grid ? grid.outerHTML.substring(0, 500) : null,
                pagerFound: !!pager,
                pagerHTML: pager ? pager.outerHTML : null
            };
        }""")
        import json
        print(json.dumps(info, indent=2))

asyncio.run(main())
