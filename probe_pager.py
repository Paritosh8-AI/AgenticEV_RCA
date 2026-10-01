import asyncio
from src.scraper.browser import BrowserSession

async def main():
    async with BrowserSession(headless=True) as page:
        await page.goto("https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails", wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
        
        tab = page.locator("//a[contains(text(), 'Cancelled')]").first
        if await tab.count() > 0:
            await tab.click()
            await page.wait_for_timeout(3000)
        
        # Check pager
        pager = page.locator(".k-pager-wrap, .k-grid-pager, .k-pager-sizes")
        print("Pager count:", await pager.count())
        for i in range(await pager.count()):
            el = pager.nth(i)
            print(f"Pager {i} text: {await el.text_content()}")
            print(f"Pager {i} html: {await el.evaluate('e => e.outerHTML[:300]')}")

        # Check options in select
        selects = page.locator(".k-pager-sizes select, select")
        print("\nSelect elements:", await selects.count())
        for i in range(await selects.count()):
            sel = selects.nth(i)
            options = await sel.locator("option").all_text_contents()
            print(f"Select {i} options: {options}")

asyncio.run(main())
