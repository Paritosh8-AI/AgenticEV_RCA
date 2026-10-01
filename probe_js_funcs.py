import sys
import asyncio
import json
from src.scraper.browser import BrowserSession

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

async def main():
    async with BrowserSession(headless=True) as page:
        target_url = "https://cms.ev-charge-network.com/Roaming/OCPIReservation/Reservation?tab=CompletedReservation"
        await page.goto(target_url, wait_until="networkidle")
        await page.wait_for_timeout(2000)

        js_funcs = await page.evaluate('''() => {
            return {
                viewStatus: typeof window.ViewStatus !== 'undefined' ? window.ViewStatus.toString() : null,
                issueHistory: typeof window.IssueHistoryPopUpReservation !== 'undefined' ? window.IssueHistoryPopUpReservation.toString() : null
            };
        }''')
        print("=== ViewStatus Definition ===")
        print(js_funcs["viewStatus"])
        print("\n=== IssueHistoryPopUpReservation Definition ===")
        print(js_funcs["issueHistory"])

asyncio.run(main())
