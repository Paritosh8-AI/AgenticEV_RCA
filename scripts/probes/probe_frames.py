import asyncio
from src.scraper.browser import BrowserSession

async def main():
    async with BrowserSession(headless=True) as page:
        url = "https://cms.ev-charge-network.com/ChargingStationManagement/AdminBookingDetails?moc=2%7CsettingsMenuItems%7CBookingDetailsInvoice%7CBookingRecords%7C106"
        await page.goto(url, wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
        
        frames = page.frames
        print(f"Total frames: {len(frames)}")
        for i, f in enumerate(frames):
            print(f"Frame {i}: name='{f.name}', url='{f.url}'")
            tables = await f.locator("table").count()
            print(f"  Tables in frame {i}: {tables}")
            if tables > 0:
                pagers = await f.locator(".k-pager-wrap, .k-grid-pager").count()
                print(f"  Pagers in frame {i}: {pagers}")
                kendo = await f.evaluate("""() => {
                    if (typeof $ !== 'undefined' && $('.k-grid').length > 0) {
                        return $('.k-grid').map(function() {
                            const g = $(this).data('kendoGrid');
                            return g ? { id: this.id, total: g.dataSource.total(), pageSize: g.dataSource.pageSize() } : null;
                        }).get();
                    }
                    return null;
                }""")
                print(f"  Kendo info in frame {i}: {kendo}")

asyncio.run(main())
