import sys
import asyncio
import json
from datetime import datetime, timedelta
from src.scraper.browser import BrowserSession

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

def get_cookie_header(storage_path="data/session_state.json") -> str:
    try:
        with open(storage_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        cookies = data.get("cookies", [])
        return "; ".join([f"{c['name']}={c['value']}" for c in cookies if "ev-charge-network.com" in c.get("domain", "")])
    except Exception:
        return ""

async def main():
    cookie_header = get_cookie_header()
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Cookie": cookie_header
    }

    async with BrowserSession(headless=True) as page:
        for label, (s_date, e_date) in [
            ("7-Day Range (Screenshot: Sep 10 - Sep 16)", ("2026-09-10", "2026-09-16")),
            ("24-Hour Range (Sep 15 - Sep 16)", ("2026-09-15", "2026-09-16"))
        ]:
            canc_url = f"https://cms.ev-charge-network.com/Roaming/OCPIReservation/LoadOcpiCancelledReservationGridViewThroughAjax?StartDate={s_date}&Enddate={e_date}&transactionTypeId=-1"
            res_canc = await page.request.post(canc_url, form={"sort": "", "page": "1", "pageSize": "1000", "group": "", "filter": ""}, headers=headers)
            d_canc = await res_canc.json()
            canc_items = d_canc.get("Data", [])

            comp_url = f"https://cms.ev-charge-network.com/Roaming/OCPIReservation/LoadOcpiCompletedReservationGridViewThroughAjax?StartDate={s_date}&Enddate={e_date}&transactionTypeId=-1"
            res_comp = await page.request.post(comp_url, form={"sort": "", "page": "1", "pageSize": "1000", "group": "", "filter": ""}, headers=headers)
            d_comp = await res_comp.json()
            comp_items = d_comp.get("Data", [])
            low_comp = [it for it in comp_items if float(it.get("KWh", 0) or 0) < 1.0]

            print(f"\n=== {label} ===")
            print(f"  Cancelled Bookings: {len(canc_items)}")
            print(f"  Completed Total: {len(comp_items)}")
            print(f"  Completed < 1 kWh: {len(low_comp)}")
            print(f"  Total Flagged Sessions: {len(canc_items) + len(low_comp)}")

asyncio.run(main())
