import asyncio
from datetime import datetime, timedelta
from src.scraper.browser import BrowserSession
from src.rca.ocpp_parser import parse_raw_ocpp_message
from src.models import OCPPLogEntry

async def fetch_logs_direct(page, entity_id: int, start_time: str, end_time: str) -> list[OCPPLogEntry]:
    url = f"https://cms.ev-charge-network.com/GetOcppLogs?entityId={entity_id}&fromDate={start_time}&toDate={end_time}"
    payload = {"sort": "", "page": "1", "pageSize": "100", "group": "", "filter": ""}
    headers = {"X-Requested-With": "XMLHttpRequest", "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"}
    
    res = await page.request.post(url, form=payload, headers=headers)
    if res.status != 200:
        return []
    try:
        data = await res.json()
        items = data.get("Data", [])
        logs = []
        for it in items:
            raw = it.get("Message", "")
            ev_name = it.get("EventName", "")
            ts = it.get("CreatedOn") or it.get("Timestamp") or ""
            parsed = parse_raw_ocpp_message(raw)
            logs.append(OCPPLogEntry(
                message_id=str(it.get("Id", "")),
                event_name=ev_name,
                event_type="Response" if "Response" in ev_name else "Request",
                timestamp=ts,
                raw_message=raw,
                message_type=parsed.get("action") or parsed.get("message_type"),
                reason=parsed.get("reason"),
                error_code=parsed.get("error_code")
            ))
        return logs
    except Exception as e:
        return []

async def main():
    async with BrowserSession(headless=True) as page:
        t1 = datetime.now()
        logs = await fetch_logs_direct(page, 5721, "2026-09-13 20:30:00", "2026-09-13 21:00:00")
        t2 = datetime.now()
        print(f"Retrieved {len(logs)} logs in {(t2 - t1).total_seconds():.2f}s")
        for l in logs[:5]:
            print(f"  [{l.timestamp}] {l.event_name} | Reason: {l.reason} | Error: {l.error_code} | Msg: {l.raw_message[:100]}")

asyncio.run(main())
