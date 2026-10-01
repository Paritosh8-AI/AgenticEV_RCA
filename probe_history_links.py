import sys
import asyncio
from src.scraper.browser import BrowserSession

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

async def main():
    async with BrowserSession(headless=True) as page:
        target_url = "https://cms.ev-charge-network.com/Roaming/OCPIReservation/Reservation?tab=CompletedReservation"
        await page.goto(target_url, wait_until="networkidle")
        await page.wait_for_timeout(3000)

        action_links = await page.evaluate('''() => {
            const cells = Array.from(document.querySelectorAll('tbody tr td:nth-child(2)'));
            return cells.slice(0, 5).map(td => td.innerHTML);
        }''')
        print("Sample Action column HTML:")
        for a in action_links:
            print("---")
            print(a)

asyncio.run(main())
