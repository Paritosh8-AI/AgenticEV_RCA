import asyncio
import json
from src.scraper.browser import BrowserSession

async def main():
    async with BrowserSession(headless=True) as page:
        cookie_header = ""
        # Let's inspect CancelledReservation tab by clicking it
        target_url = "https://cms.ev-charge-network.com/Roaming/OCPIReservation/Reservation?moc=2%7CsettingsMenuItems%7CNA%7CReservation%7C378&tab=CompletedReservation"
        await page.goto(target_url, wait_until="networkidle")
        await page.wait_for_timeout(2000)

        captured = []
        async def on_req(req):
            if "Ajax" in req.url or "Reservation" in req.url:
                captured.append({"url": req.url, "method": req.method, "post": req.post_data})

        page.on("request", on_req)

        # Click Cancelled tab
        print("Clicking Cancelled tab...")
        canc_tab = page.locator("#CancelledReservation, a:has-text('Cancelled')").first
        if await canc_tab.count() > 0:
            await canc_tab.click()
            await page.wait_for_timeout(4000)

        print("\nCaptured Requests after clicking Cancelled:")
        for c in captured:
            print(c["url"], "POST:", c["post"])

        # Also let's inspect the DOM of Cancelled table
        canc_info = await page.evaluate('''() => {
            const headers = Array.from(document.querySelectorAll('.k-grid-header th, table thead th')).map(th => th.innerText.trim());
            const rows = Array.from(document.querySelectorAll('.k-grid-content tr, table tbody tr')).slice(0, 3).map(tr => {
                return Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim());
            });
            return { headers, rows };
        }''')
        print("\nCancelled Table Headers:")
        print(canc_info["headers"])
        print("\nCancelled Sample Rows:")
        print(canc_info["rows"])

asyncio.run(main())
