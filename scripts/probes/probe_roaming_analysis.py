import sys
import asyncio
import json
from collections import Counter
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

    start_date = "2026-09-10"
    end_date = "2026-09-16"

    async with BrowserSession(headless=True) as page:
        # Cancelled
        canc_url = f"https://cms.ev-charge-network.com/Roaming/OCPIReservation/LoadOcpiCancelledReservationGridViewThroughAjax?StartDate={start_date}&Enddate={end_date}&transactionTypeId=-1"
        res_canc = await page.request.post(canc_url, form={"sort": "", "page": "1", "pageSize": "1000", "group": "", "filter": ""}, headers=headers)
        canc_items = (await res_canc.json()).get("Data", [])

        # Completed
        comp_url = f"https://cms.ev-charge-network.com/Roaming/OCPIReservation/LoadOcpiCompletedReservationGridViewThroughAjax?StartDate={start_date}&Enddate={end_date}&transactionTypeId=-1"
        res_comp = await page.request.post(comp_url, form={"sort": "", "page": "1", "pageSize": "1000", "group": "", "filter": ""}, headers=headers)
        comp_items = (await res_comp.json()).get("Data", [])
        low_comp = [it for it in comp_items if float(it.get("KWh", 0) or 0) < 1.0]

        print(f"Total Cancelled fetched: {len(canc_items)}")
        print(f"Total Low Consumption (<1 kWh) fetched: {len(low_comp)}")

        # Clean SchedularAction for Cancelled
        # Usually it has timestamps: "Cancelled by Scheduler on 2026-09-14 09:56 PM" -> base: "Cancelled by Scheduler"
        def normalize_action(act: str) -> str:
            if not act or not act.strip():
                return "Unspecified / No Action Recorded"
            a = act.strip()
            import re
            # Strip timestamps like "2026-09-16 14:20:58" or "on 2026-09-15 06:48 PM"
            a_clean = re.sub(r'\s+on\s+\d{4}-\d{2}-\d{2}\s+.*', '', a, flags=re.IGNORECASE)
            a_clean = re.sub(r'\s+\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}(:\d{2})?(\s+[AP]M)?', '', a_clean, flags=re.IGNORECASE)
            return a_clean.strip()

        canc_reasons = Counter()
        for it in canc_items:
            norm = normalize_action(it.get("SchedularAction"))
            canc_reasons[norm] += 1

        print("\n=== CANCELLED REASONS (Normalized) ===")
        for k, v in canc_reasons.most_common():
            print(f"  {k}: {v} ({(v/len(canc_items)*100):.1f}%)")

        # Now let's see Low Consumption completed sessions:
        # What fields exist on Low Consumption completed sessions?
        comp_reasons = Counter()
        for it in low_comp:
            sched = normalize_action(it.get("SchedularAction"))
            dur = it.get("TimeDuration")
            kwh = float(it.get("KWh", 0) or 0)
            if kwh == 0.0:
                comp_reasons["Zero Energy Consumed (0.0 kWh) - Premature Termination"] += 1
            else:
                comp_reasons[f"Low Energy Transfer (<1 kWh: {kwh:.2f} kWh)"] += 1

        print("\n=== LOW CONSUMPTION REASONS ===")
        for k, v in comp_reasons.most_common():
            print(f"  {k}: {v} ({(v/len(low_comp)*100):.1f}%)")

        # Breakdown by Party ID
        print("\n=== CANCELLED BREAKDOWN BY PARTY ID ===")
        party_canc = Counter()
        for it in canc_items:
            p = f"{it.get('PartyId') or 'Unknown'} ({it.get('PartyName') or 'Unknown'})"
            party_canc[p] += 1
        for k, v in party_canc.most_common(10):
            print(f"  {k}: {v}")

        print("\n=== LOW CONSUMPTION BREAKDOWN BY PARTY ID ===")
        party_comp = Counter()
        for it in low_comp:
            p = f"{it.get('PartyId') or 'Unknown'} ({it.get('PartyName') or 'Unknown'})"
            party_comp[p] += 1
        for k, v in party_comp.most_common(10):
            print(f"  {k}: {v}")

        # Breakdown by Manufacturer
        print("\n=== CANCELLED BREAKDOWN BY VEHICLE MANUFACTURER ===")
        mfg_canc = Counter()
        for it in canc_items:
            m = it.get('ManufacturerName') or 'Unknown'
            mfg_canc[m] += 1
        for k, v in mfg_canc.most_common(10):
            print(f"  {k}: {v}")

        print("\n=== LOW CONSUMPTION BREAKDOWN BY VEHICLE MANUFACTURER ===")
        mfg_comp = Counter()
        for it in low_comp:
            m = it.get('ManufacturerName') or 'Unknown'
            mfg_comp[m] += 1
        for k, v in mfg_comp.most_common(10):
            print(f"  {k}: {v}")

        # Breakdown by Station Name
        print("\n=== TOP 10 AFFECTED STATIONS (CANCELLED) ===")
        st_canc = Counter()
        for it in canc_items:
            st = it.get('StationName') or 'Unknown'
            st_canc[st] += 1
        for k, v in st_canc.most_common(10):
            print(f"  {k}: {v}")

        print("\n=== TOP 10 AFFECTED STATIONS (LOW CONSUMPTION) ===")
        st_comp = Counter()
        for it in low_comp:
            st = it.get('StationName') or 'Unknown'
            st_comp[st] += 1
        for k, v in st_comp.most_common(10):
            print(f"  {k}: {v}")

asyncio.run(main())
