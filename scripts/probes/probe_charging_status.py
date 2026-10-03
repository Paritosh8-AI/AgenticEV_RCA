import sys
import asyncio
from src.scraper.browser import BrowserSession

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

async def main():
    async with BrowserSession(headless=True) as page:
        for bid in ["151164", "151166"]:
            url = f"https://cms.ev-charge-network.com/Roaming/OCPIReservation/GetChargingStatus?bookingId={bid}"
            res = await page.request.get(url)
            print(f"\n=== GetChargingStatus for {bid} (Status: {res.status}) ===")
            text = await res.text()
            print("Content Length:", len(text))
            print("Snippet (first 1000 chars):")
            print(text[:1000])

asyncio.run(main())
