import asyncio
from src.scraper.browser import BrowserSession
from src.scraper.bookings import fetch_bookings_list

async def test():
    async with BrowserSession(headless=True) as page:
        # 1. Fetch Cancelled for last 24h
        print("Fetching Cancelled bookings for 2026-09-12 to 2026-09-13...")
        cancelled = await fetch_bookings_list(
            page=page,
            start_date="2026-09-12",
            end_date="2026-09-13",
            tab="Cancelled",
            max_rows=100
        )
        print(f"Cancelled retrieved: {len(cancelled)}")
        for r in cancelled[:5]:
            print(f"  [{r.booking_id}] Time: {r.booking_start_time} | Charger: {r.station_organization} | Conn: {r.connector_id}")

        # 2. Fetch Completed for last 24h
        print("\nFetching Completed bookings for 2026-09-12 to 2026-09-13...")
        completed = await fetch_bookings_list(
            page=page,
            start_date="2026-09-12",
            end_date="2026-09-13",
            tab="Completed",
            max_rows=100
        )
        print(f"Completed retrieved: {len(completed)}")
        for r in completed[:5]:
            print(f"  [{r.booking_id}] Time: {r.booking_start_time} | Charger: {r.station_organization} | Conn: {r.connector_id}")

asyncio.run(test())
