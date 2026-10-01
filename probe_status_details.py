import sys
import asyncio
from bs4 import BeautifulSoup
from src.scraper.browser import BrowserSession

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

async def main():
    async with BrowserSession(headless=True) as page:
        for bid in ["151164", "151166"]:
            url = f"https://cms.ev-charge-network.com/Roaming/OCPIReservation/GetChargingStatus?bookingId={bid}"
            res = await page.request.get(url)
            html = await res.text()
            soup = BeautifulSoup(html, "html.parser")
            
            print(f"\n==========================================")
            print(f"BOOKING ID: {bid}")
            print(f"==========================================")
            
            # Find all text in labels, cards, dt/dd, table rows, badges
            labels = []
            for el in soup.find_all(["label", "dt", "dd", "th", "td", "span", "p", "h4", "h5"]):
                txt = el.get_text(strip=True)
                if txt and len(txt) < 100:
                    labels.append(txt)
            
            # Filter and display interesting labels
            interesting = []
            for t in labels:
                low = t.lower()
                if any(k in low for k in ["stop", "reason", "status", "error", "fault", "cancel", "ocpi", "cpo", "session", "kwh", "soc", "schedule", "abort", "disconnect"]):
                    interesting.append(t)
            
            print("Interesting keywords found:")
            for item in set(interesting[:30]):
                print(f"  - {item}")
                
            # Look for tables or cards
            print("\nCard / Table text blocks:")
            for card in soup.find_all(class_=["card", "box", "panel", "modal-body", "container-fluid"]):
                card_text = card.get_text(separator=" | ", strip=True)
                if len(card_text) > 20 and len(card_text) < 1000:
                    print("CARD:", card_text[:400])

asyncio.run(main())
