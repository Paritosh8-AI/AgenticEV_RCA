import sys
import os
import asyncio
import json
from datetime import datetime, timedelta
from src.scraper.browser import BrowserSession

# Fix standard output encoding for Windows
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

    start_date = "2026-09-10"
    end_date = "2026-09-16"

    async with BrowserSession(headless=True) as page:
        # Pre-warm
        await page.goto("https://cms.ev-charge-network.com/Roaming/OCPIReservation/Reservation?tab=CompletedReservation", wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)

        # 1. Fetch Cancelled sample
        canc_url = f"https://cms.ev-charge-network.com/Roaming/OCPIReservation/LoadOcpiCancelledReservationGridViewThroughAjax?StartDate={start_date}&Enddate={end_date}&transactionTypeId=-1"
        payload = {"sort": "", "page": "1", "pageSize": "10", "group": "", "filter": ""}
        
        print(f"Fetching Cancelled from {canc_url}...")
        res_canc = await page.request.post(canc_url, form=payload, headers=headers)
        data_canc = await res_canc.json()
        canc_items = data_canc.get("Data", [])
        print(f"Cancelled Total: {data_canc.get('Total')} | Retrieved: {len(canc_items)}")
        if canc_items:
            print("\n=== SAMPLE CANCELLED ROW (Keys & Values) ===")
            for k, v in canc_items[0].items():
                print(f"  {k}: {repr(v)}")

        # 2. Fetch Completed sample
        comp_url = f"https://cms.ev-charge-network.com/Roaming/OCPIReservation/LoadOcpiCompletedReservationGridViewThroughAjax?StartDate={start_date}&Enddate={end_date}&transactionTypeId=-1"
        print(f"\nFetching Completed from {comp_url}...")
        res_comp = await page.request.post(comp_url, form=payload, headers=headers)
        data_comp = await res_comp.json()
        comp_items = data_comp.get("Data", [])
        print(f"Completed Total: {data_comp.get('Total')} | Retrieved: {len(comp_items)}")
        if comp_items:
            print("\n=== SAMPLE COMPLETED ROW (Keys & Values) ===")
            for k, v in comp_items[0].items():
                print(f"  {k}: {repr(v)}")

            # Check low energy
            print("\nCompleted items energy values:")
            for it in comp_items:
                for ek in ["Kwh", "EnergyConsumed", "UnitsConsumed", "Energy", "kWh"]:
                    if ek in it:
                        print(f"  ID: {it.get('ChargingStationBookingId') or it.get('BookingId') or it.get('BookingID')} | {ek}: {it[ek]}")

asyncio.run(main())
