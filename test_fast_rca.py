import asyncio
import json
from datetime import datetime, timedelta
from src.scraper.browser import BrowserSession
from src.models import OCPPLogEntry, BookingDetails
from src.rca.ocpp_parser import parse_ocpp_message
from src.rca.engine import analyze_booking_rca
from generate_excel_rca import map_to_cs_rca_text

async def fetch_logs_for_item(page, sem, item):
    entity_id = item.get("ChargingStationChargerId")
    charger_code = item.get("ChargerCode") or ""
    b_time_str = item.get("BookingInTime") or item.get("DateForBooking") or ""
    
    dt = None
    for fmt in ["%Y-%m-%d %I:%M %p", "%Y-%m-%d %I:%M:%S %p", "%Y-%m-%d"]:
        try:
            dt = datetime.strptime(b_time_str.strip(), fmt)
            break
        except:
            continue
    if not dt:
        dt = datetime.now()
        
    start_time = (dt - timedelta(minutes=15)).strftime("%Y-%m-%d %H:%M:00")
    end_time = (dt + timedelta(minutes=15)).strftime("%Y-%m-%d %H:%M:00")
    
    url = f"https://cms.ev-charge-network.com/GetOcppLogs?entityId={entity_id}&fromDate={start_time}&toDate={end_time}"
    payload = {"sort": "", "page": "1", "pageSize": "100", "group": "", "filter": ""}
    headers = {"X-Requested-With": "XMLHttpRequest", "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
    
    async with sem:
        try:
            res = await page.request.post(url, form=payload, headers=headers)
            if res.status != 200:
                return []
            data = await res.json()
            items = data.get("Data", [])
            logs = []
            for it in items:
                raw = it.get("Message", "")
                ev_name = it.get("EventName", "")
                ts = it.get("CreatedOn") or it.get("Timestamp") or ""
                parsed_entry = parse_ocpp_message(
                    raw_message=raw,
                    event_name=ev_name,
                    timestamp=ts,
                    message_id_fallback=str(it.get("Id", ""))
                )
                logs.append(parsed_entry)
            return logs
        except Exception:
            return []

async def main():
    async with BrowserSession(headless=True) as page:
        payload = {"sort": "", "page": "1", "pageSize": "1000", "group": "", "filter": ""}
        headers = {"X-Requested-With": "XMLHttpRequest", "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
        url = "https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails/CancelledBookingDataThroughAjax?StationId=&StartDate=2026-09-12&Enddate=2026-09-13&stateId=&cityId=&transactionTypeId=-1&Id=0&vin=&chargerCode=&user=&vehicleNumber="
        res = await page.request.post(url, form=payload, headers=headers)
        data = await res.json()
        all_items = data.get("Data", [])
        
        test_items = all_items[:10]
        print(f"Testing fast RCA on {len(test_items)} items...")
        
        sem = asyncio.Semaphore(5)
        tasks = [fetch_logs_for_item(page, sem, it) for it in test_items]
        
        t1 = datetime.now()
        logs_list = await asyncio.gather(*tasks)
        t2 = datetime.now()
        print(f"Fetched logs for 10 items in {(t2 - t1).total_seconds():.2f}s!")
        
        for it, logs in zip(test_items, logs_list):
            bid = it.get("ChargingStationBookingId")
            b_details = BookingDetails(
                booking_id=str(bid),
                station_name=it.get("ChargingStationName"),
                charger_code=it.get("ChargerCode"),
                connector_sequence_id=it.get("ConnectorId"),
                booking_in_time=it.get("BookingInTime"),
                booking_date=it.get("DateForBooking")
            )
            rca = analyze_booking_rca(b_details, logs)
            cs_text = map_to_cs_rca_text(rca, logs)
            print(f"Booking #{bid} | {it.get('ChargingStationName')} ({it.get('ChargerCode')}) -> {cs_text}")

asyncio.run(main())
