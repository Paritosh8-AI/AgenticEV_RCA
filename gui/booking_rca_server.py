"""
ElectreeFi Booking RCA Deep Dive Studio - Dedicated Backend Server
Provides specialized HTTP API endpoints and static UI serving exclusively
for Single Booking RCA & Telemetry Investigation.
"""

import os
import sys
import json
import pathlib
import datetime
import subprocess
from http import HTTPStatus
from http.server import HTTPServer, SimpleHTTPRequestHandler
import urllib.parse

if getattr(sys, "frozen", False):
    exe_dir = pathlib.Path(sys.executable).resolve().parent
    if (exe_dir / "data").exists():
        BASE_DIR = exe_dir
    elif (exe_dir.parent / "data").exists():
        BASE_DIR = exe_dir.parent
    else:
        BASE_DIR = exe_dir

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass and (pathlib.Path(meipass) / "gui" / "static_booking_rca").exists():
        STATIC_DIR = pathlib.Path(meipass) / "gui" / "static_booking_rca"
    elif (BASE_DIR / "gui" / "static_booking_rca").exists():
        STATIC_DIR = BASE_DIR / "gui" / "static_booking_rca"
    else:
        STATIC_DIR = pathlib.Path(__file__).resolve().parent / "static_booking_rca"
else:
    BASE_DIR = pathlib.Path(__file__).resolve().parent.parent
    STATIC_DIR = pathlib.Path(__file__).resolve().parent / "static_booking_rca"

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

VENV_PYTHON = BASE_DIR / ".venv" / "Scripts" / "python.exe"
PYTHON_EXE = str(VENV_PYTHON) if VENV_PYTHON.exists() else sys.executable


def check_session_detailed() -> dict:
    """Checks session state and performs fast validation."""
    session_file = BASE_DIR / "data" / "session_state.json"
    if not session_file.exists():
        return {
            "has_session": False,
            "status_text": "No Session File Found",
            "modified": None,
            "is_valid": False
        }
    
    try:
        stat = session_file.stat()
        mod_time = datetime.datetime.fromtimestamp(stat.st_mtime)
        mod_str = mod_time.strftime("%d-%b-%Y %H:%M")
        
        # Test validity
        from src.auth.session_manager import is_session_valid_sync
        valid = is_session_valid_sync()
        status_text = "Session Active & Verified" if valid else "Session Token Expired"
        
        return {
            "has_session": True,
            "status_text": status_text,
            "modified": mod_str,
            "is_valid": valid
        }
    except Exception as e:
        return {
            "has_session": True,
            "status_text": f"Session Check Error: {str(e)}",
            "modified": None,
            "is_valid": False
        }


class BookingRcaHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_DIR), **kwargs)

    def log_message(self, format, *args):
        # Suppress spammy HTTP request logs from console
        pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/api/health":
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({
                "status": "ok",
                "app": "ElectreeFi Booking RCA Deep Dive Studio",
                "version": "2.0.0"
            }).encode("utf-8"))
            return

        elif path == "/api/session-status":
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            status_data = check_session_detailed()
            self.wfile.write(json.dumps(status_data).encode("utf-8"))
            return

        # Serve static UI index.html on root
        if path in ["", "/"]:
            self.path = "/index.html"

        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"

        try:
            payload = json.loads(body)
        except Exception:
            payload = {}

        if path == "/api/search-booking":
            from src.rca.booking_inspector import investigate_booking, clear_investigation_cache

            booking_id = str(payload.get("booking_id", "")).strip()
            start_date = payload.get("start_date") or None
            end_date = payload.get("end_date") or None
            bypass_cache = bool(payload.get("bypass_cache", False))

            if bypass_cache:
                clear_investigation_cache()

            try:
                if start_date and end_date:
                    result = investigate_booking(booking_id, start_date=start_date, end_date=end_date)
                elif start_date:
                    result = investigate_booking(booking_id, start_date=start_date)
                else:
                    result = investigate_booking(booking_id)
            except Exception as e:
                result = {
                    "success": False,
                    "found": False,
                    "message": f"Server error during investigation: {str(e)}"
                }

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(result).encode("utf-8"))
            return

        elif path == "/api/clear-cache":
            from src.rca.booking_inspector import clear_investigation_cache
            clear_investigation_cache()

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({
                "success": True,
                "message": "Investigation cache purged successfully."
            }).encode("utf-8"))
            return

        elif path == "/api/trigger-login":
            try:
                # Launch interactive login in background
                cmd = [PYTHON_EXE, "-m", "src.cli", "login", "--force"]
                subprocess.Popen(cmd, cwd=str(BASE_DIR))
                success = True
                msg = "Launched interactive CMS login window. Please complete CAPTCHA & OTP in the opened browser."
            except Exception as e:
                success = False
                msg = f"Failed to launch login: {str(e)}"

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"success": success, "message": msg}).encode("utf-8"))
            return

        elif path == "/api/open-folder":
            try:
                os.startfile(str(BASE_DIR))
                success = True
                msg = "Opened folder in Explorer"
            except Exception as e:
                success = False
                msg = str(e)

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"success": success, "message": msg}).encode("utf-8"))
            return

        self.send_error(HTTPStatus.NOT_FOUND, "API endpoint not found")


def start_server(port: int = 58220):
    server = HTTPServer(("127.0.0.1", port), BookingRcaHandler)
    print(f"ElectreeFi Booking RCA Studio Server started on http://127.0.0.1:{port}")
    return server


if __name__ == "__main__":
    port = 58220
    server = start_server(port)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down Booking RCA Studio...")
        server.shutdown()
