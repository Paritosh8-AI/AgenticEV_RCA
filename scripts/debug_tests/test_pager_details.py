import asyncio
from src.scraper.browser import BrowserSession

async def main():
    async with BrowserSession(headless=True) as page:
        url = "https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails"
        await page.goto(url, wait_until="domcontentloaded")
        
        # Wait for Cancelled tab
        tab = page.locator("//a[text()='Cancelled' or contains(text(), 'Cancelled')]").first
        await tab.wait_for(state="visible", timeout=15000)
        await tab.click()
        
        # Wait for table to load
        table = page.locator("table.k-selectable, table:has(tbody tr td)").first
        await table.wait_for(state="visible", timeout=20000)
        await page.wait_for_timeout(2000)

        rows = await table.locator("tbody tr").count()
        print(f"Initial visible rows on Cancelled tab: {rows}")

        # Now inspect the pager below the table!
        pager_info = await page.evaluate("""() => {
            // Check all elements around table or at bottom
            const pagers = document.querySelectorAll('.k-pager-wrap, .k-grid-pager');
            return Array.from(pagers).map(p => {
                const select = p.querySelector('select');
                const dropdown = p.querySelector('.k-dropdown, [data-role=\"dropdownlist\"]');
                const infoText = p.querySelector('.k-pager-info');
                return {
                    text: p.innerText,
                    hasSelect: !!select,
                    selectHTML: select ? select.outerHTML : null,
                    hasDropdown: !!dropdown,
                    dropdownHTML: dropdown ? dropdown.outerHTML : null,
                    infoText: infoText ? infoText.innerText : null
                };
            });
        }""")
        import json
        print("Pager info:", json.dumps(pager_info, indent=2))

asyncio.run(main())
