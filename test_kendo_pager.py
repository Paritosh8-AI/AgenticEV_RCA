import asyncio
from src.scraper.browser import BrowserSession

async def main():
    async with BrowserSession(headless=True) as page:
        print("Navigating to AdminBookingDetails...")
        await page.goto("https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails", wait_until="load")
        await page.wait_for_timeout(3000)

        # Click Cancelled tab
        print("Clicking Cancelled tab...")
        cancelled_tab = page.locator("//a[text()='Cancelled' or contains(text(), 'Cancelled')]").first
        await cancelled_tab.click()
        
        # Wait for pager info or grid
        print("Waiting for .k-pager-info...")
        pager_info = page.locator(".k-pager-info")
        await pager_info.wait_for(state="visible", timeout=30000)
        text = await pager_info.text_content()
        print(f"Current pager info text: {text}")

        # Check pager sizes dropdown
        pager_sizes = page.locator(".k-pager-sizes")
        print(f"Pager sizes HTML: {await pager_sizes.evaluate('e => e.outerHTML')}")

        # Try selecting 1000
        select_el = page.locator(".k-pager-sizes select")
        if await select_el.count() > 0:
            print("Selecting 1000 on select element...")
            await select_el.select_option("1000")
        else:
            print("Clicking k-dropdown...")
            dropdown = page.locator(".k-pager-sizes .k-dropdown, .k-pager-sizes span.k-widget")
            await dropdown.click()
            await page.wait_for_timeout(1000)
            opt_1000 = page.locator("//li[text()='1000' or contains(text(), '1000')]")
            await opt_1000.click()

        print("Waiting for grid reload...")
        await page.wait_for_timeout(5000)
        new_text = await pager_info.text_content()
        print(f"New pager info text after selecting 1000: {new_text}")

        # Count table rows
        rows = await page.locator("table.k-selectable tbody tr, .k-grid-content tbody tr").count()
        print(f"Total rows now rendered in DOM: {rows}")

asyncio.run(main())
