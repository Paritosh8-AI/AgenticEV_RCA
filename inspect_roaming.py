import asyncio
import json
from src.scraper.browser import BrowserSession

async def main():
    async with BrowserSession(headless=True) as page:
        captured_requests = []
        captured_responses = []

        async def on_request(req):
            if any(k in req.url for k in ['Reservation', 'OCPI', 'Roaming', 'Ajax', 'Data']):
                captured_requests.append({
                    'url': req.url,
                    'method': req.method,
                    'post_data': req.post_data
                })

        async def on_response(res):
            if any(k in res.url for k in ['Reservation', 'OCPI', 'Roaming', 'Ajax', 'Data']):
                try:
                    data = await res.json()
                    captured_responses.append({
                        'url': res.url,
                        'status': res.status,
                        'json_snippet': str(data)[:500] if data else None,
                        'keys': list(data.keys()) if isinstance(data, dict) else (f'list of {len(data)}' if isinstance(data, list) else None)
                    })
                except:
                    pass

        page.on('request', on_request)
        page.on('response', on_response)

        target_url = 'https://cms.ev-charge-network.com/Roaming/OCPIReservation/Reservation?moc=2%7CsettingsMenuItems%7CNA%7CReservation%7C378&tab=CompletedReservation'
        print(f'Navigating to {target_url}...')
        await page.goto(target_url, wait_until='networkidle')
        await page.wait_for_timeout(3000)

        # Inspect DOM elements
        info = await page.evaluate('''() => {
            // Find tabs
            const tabs = Array.from(document.querySelectorAll('.nav-tabs a, ul.k-tabstrip-items li, [role="tab"], a[data-toggle="tab"], .tab, button.tab')).map(el => ({
                text: el.innerText.trim(),
                href: el.getAttribute('href'),
                id: el.id,
                class: el.className
            }));

            // Find dropdowns (like Party ID)
            const selects = Array.from(document.querySelectorAll('select')).map(s => ({
                name: s.name,
                id: s.id,
                options: Array.from(s.options).map(o => ({ value: o.value, text: o.text }))
            }));

            // Find grid headers
            const headers = Array.from(document.querySelectorAll('.k-grid-header th, table thead th')).map(th => th.innerText.trim());

            // Sample row data
            const rows = Array.from(document.querySelectorAll('.k-grid-content tr, table tbody tr')).slice(0, 3).map(tr => {
                return Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim());
            });

            return { tabs, selects, headers, rows };
        }''')

        print('\n=== TABS ===')
        print(json.dumps(info['tabs'], indent=2))

        print('\n=== SELECTS / DROPDOWNS ===')
        print(json.dumps(info['selects'], indent=2))

        print('\n=== GRID HEADERS ===')
        print(json.dumps(info['headers'], indent=2))

        print('\n=== SAMPLE ROWS ===')
        print(json.dumps(info['rows'], indent=2))

        print('\n=== CAPTURED AJAX REQUESTS ===')
        for r in captured_requests:
            print(f"{r['method']} {r['url']} [PostData: {r['post_data']}]")

        print('\n=== CAPTURED AJAX RESPONSES ===')
        for r in captured_responses:
            print(f"{r['url']} -> {r['status']} [Keys: {r['keys']}]")

asyncio.run(main())
