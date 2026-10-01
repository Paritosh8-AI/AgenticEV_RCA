import asyncio
from src.scraper.browser import BrowserSession

async def main():
    async with BrowserSession(headless=True) as page:
        await page.goto("https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails", wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
        
        # Click Cancelled tab
        await page.locator("//a[text()='Cancelled' or contains(text(), 'Cancelled')]").first.click()
        print("Clicked Cancelled. Waiting 4s...")
        await page.wait_for_timeout(4000)

        # Inspect Kendo Grid DataSource
        res = await page.evaluate("""() => {
            // Find all Kendo grids on page
            const grids = $('.k-grid');
            const info = [];
            grids.each(function() {
                const grid = $(this).data('kendoGrid');
                if (grid) {
                    info.push({
                        id: this.id,
                        total: grid.dataSource.total(),
                        pageSize: grid.dataSource.pageSize(),
                        totalPages: grid.dataSource.totalPages()
                    });
                }
            });
            return info;
        }""")
        print("Kendo Grids info:", res)

asyncio.run(main())
