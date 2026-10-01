import asyncio
from src.scraper.browser import BrowserSession
from src.scraper.bookings import fetch_bookings_list

async def main():
    async with BrowserSession(headless=True) as page:
        # fetch_bookings_list navigates and clicks Cancelled
        await fetch_bookings_list(page, tab="Cancelled", max_rows=5)
        
        info = await page.evaluate("""() => {
            const table = document.querySelector('table.k-selectable, table:has(tbody tr td)');
            const gridEl = table ? table.closest('.k-grid') : null;
            const pagerEl = document.querySelector('.k-pager-wrap, .k-grid-pager');
            
            let kendoGridFound = false;
            let options = [];
            let total = 0;
            let currentSize = 0;
            
            if (gridEl && typeof $(gridEl).data('kendoGrid') !== 'undefined') {
                const g = $(gridEl).data('kendoGrid');
                kendoGridFound = true;
                total = g.dataSource.total();
                currentSize = g.dataSource.pageSize();
            }

            // Also check all select elements inside pager
            const pagerSelect = pagerEl ? pagerEl.querySelector('select') : null;
            if (pagerSelect) {
                options = Array.from(pagerSelect.options).map(o => ({ value: o.value, text: o.text }));
            }

            return {
                gridId: gridEl ? gridEl.id : null,
                kendoGridFound: kendoGridFound,
                total: total,
                currentSize: currentSize,
                pagerHtml: pagerEl ? pagerEl.outerHTML : null,
                selectOptions: options
            };
        }""")
        
        print("Grid inspection info:")
        print(f"  gridId: {info['gridId']}")
        print(f"  kendoGridFound: {info['kendoGridFound']}")
        print(f"  total: {info['total']}")
        print(f"  currentSize: {info['currentSize']}")
        print(f"  selectOptions: {info['selectOptions']}")
        if info['pagerHtml']:
            print(f"  pagerHtml (snippet): {info['pagerHtml'][:300]}")

asyncio.run(main())
