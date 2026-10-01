import asyncio
from src.scraper.browser import BrowserSession

async def main():
    async with BrowserSession(headless=True) as page:
        await page.goto("https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails", wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)

        # Click Cancelled tab
        tab_elem = page.locator("//a[contains(text(), 'Cancelled')]").first
        if await tab_elem.count() > 0:
            async with page.expect_response(lambda r: "CancelledBookingDataThroughAjax" in r.url, timeout=15000):
                await tab_elem.click()
            await page.wait_for_timeout(1500)
            print("Switched to tab Cancelled")

        # Check rows before
        data_table = page.locator("table.k-selectable, table:has(tbody tr td)").first
        rows_before = await data_table.locator("tbody tr").count()
        print(f"Rows before changing pageSize: {rows_before}")

        # Check what Kendo grid IDs exist on the page
        grid_info = await page.evaluate("""() => {
            const results = [];
            document.querySelectorAll('[data-role="grid"]').forEach(el => {
                results.push({
                    id: el.id,
                    className: el.className,
                    hasKendo: typeof $(el).data('kendoGrid') !== 'undefined'
                });
            });
            // Also check all elements with class k-grid
            document.querySelectorAll('.k-grid').forEach(el => {
                results.push({
                    id: el.id,
                    className: el.className,
                    hasKendo: typeof $(el).data('kendoGrid') !== 'undefined'
                });
            });
            return results;
        }""")
        print("Grid elements found:", grid_info)

        # Try changing page size via Kendo API or dropdown
        change_res = await page.evaluate("""() => {
            // Check all jQuery kendo grids
            let changed = false;
            $('.k-grid').each(function() {
                const g = $(this).data('kendoGrid');
                if (g && g.dataSource) {
                    g.dataSource.pageSize(1000);
                    changed = true;
                }
            });
            return changed;
        }""")
        print(f"Changed page size via Kendo API: {change_res}")
        await page.wait_for_timeout(3000)

        rows_after = await data_table.locator("tbody tr").count()
        print(f"Rows after changing pageSize via API: {rows_after}")

        # If that didn't work, let's look at the dropdown in .k-pager-sizes
        pager_sizes = page.locator(".k-pager-sizes, .k-dropdown:has-text('30'), .k-dropdown:has-text('10')")
        print("Pager sizes count:", await pager_sizes.count())
        for i in range(await pager_sizes.count()):
            print(f"Pager size {i}: {await pager_sizes.nth(i).text_content()}")

asyncio.run(main())
