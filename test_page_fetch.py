import asyncio
from src.scraper.browser import BrowserSession

async def main():
    async with BrowserSession(headless=True) as page:
        await page.goto("https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails", wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
        
        # Test fetch from within the page context
        res = await page.evaluate("""async () => {
            const formData = new URLSearchParams();
            formData.append('sort', '');
            formData.append('page', '1');
            formData.append('pageSize', '1000');
            formData.append('group', '');
            formData.append('filter', '');
            
            const url = '/ChargingStationManagement/AdminBookingDetails/CancelledBookingDataThroughAjax?StationId=&StartDate=2026-09-12&Enddate=2026-09-13&stateId=&cityId=&transactionTypeId=-1&Id=0&vin=&chargerCode=&user=&vehicleNumber=';
            
            const r = await fetch(url, {
                method: 'POST',
                headers: {
                    'X-Requested-With': 'XMLHttpRequest',
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
                },
                body: formData.toString()
            });
            
            const text = await r.text();
            return {
                status: r.status,
                text: text.substring(0, 300)
            };
        }""")
        
        print("Fetch Status:", res["status"])
        items = res["data"].get("Data", [])
        print("Items returned:", len(items))

asyncio.run(main())
