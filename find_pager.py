import asyncio
from src.scraper.browser import BrowserSession

async def main():
    async with BrowserSession(headless=True) as page:
        await page.goto("https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails", wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
        
        # Click Cancelled tab
        await page.locator("//a[text()='Cancelled' or contains(text(), 'Cancelled')]").first.click()
        print("Clicked Cancelled tab. Waiting 5s...")
        await page.wait_for_timeout(5000)
        
        # Look for pager
        pagers = await page.locator(".k-pager-wrap").all()
        print(f"Found {len(pagers)} pager elements with .k-pager-wrap")
        for i, p in enumerate(pagers):
            text = await p.text_content()
            html = await p.evaluate("e => e.outerHTML")
            print(f"\n--- Pager {i} Text ---:\n{text}")
            print(f"--- Pager {i} HTML (first 400 chars) ---:\n{html[:400]}")

        # Also find all dropdowns or selects
        selects = await page.locator("select").all()
        print(f"\nFound {len(selects)} select elements:")
        for s in selects:
            opts = await s.locator("option").all_text_contents()
            val = await s.evaluate("e => e.value")
            name = await s.evaluate("e => e.name || e.id || e.className")
            print(f"Select '{name}': current='{val}', options={opts}")

asyncio.run(main())
