"""
ElectreeFi Automation Studio - Backend Application Server
Provides HTTP API endpoints, task orchestration, real-time log streaming, and desktop integration.
"""

import os
import sys
import re
import json
import time
import queue
import base64
import pathlib
import datetime
import threading
import subprocess
from http import HTTPStatus
from http.server import HTTPServer, SimpleHTTPRequestHandler
import urllib.parse

if getattr(sys, "frozen", False):
    BASE_DIR = pathlib.Path(sys.executable).resolve().parent
    meipass = getattr(sys, "_MEIPASS", None)
    if (BASE_DIR / "gui" / "static").exists():
        STATIC_DIR = BASE_DIR / "gui" / "static"
    elif meipass and (pathlib.Path(meipass) / "gui" / "static").exists():
        STATIC_DIR = pathlib.Path(meipass) / "gui" / "static"
    else:
        STATIC_DIR = pathlib.Path(__file__).resolve().parent / "static"
else:
    BASE_DIR = pathlib.Path(__file__).resolve().parent.parent
    STATIC_DIR = pathlib.Path(__file__).resolve().parent / "static"

VENV_PYTHON = BASE_DIR / ".venv" / "Scripts" / "python.exe"
PYTHON_EXE = str(VENV_PYTHON) if VENV_PYTHON.exists() else sys.executable

# Report files tracked by the hub
TRACKED_FILES = [
    {
        "id": "regular_excel",
        "name": "ElectreeFi_Cancelled_Bookings_RCA.xlsx",
        "title": "Cancelled Bookings RCA (Excel)",
        "type": "excel",
        "category": "Regular Bookings"
    },
    {
        "id": "regular_excel_all",
        "name": "ElectreeFi_Cancelled_Bookings_RCA_All_Records.xlsx",
        "title": "Cancelled Bookings RCA - All Records (Excel)",
        "type": "excel",
        "category": "Regular Bookings"
    },
    {
        "id": "regular_word",
        "name": "ElectreeFi_24Hour_RCA_Report.docx",
        "title": "24-Hour Executive RCA Report (Word)",
        "type": "word",
        "category": "Regular Bookings"
    },
    {
        "id": "roaming_excel",
        "name": "ElectreeFi_Roaming_Reservation_RCA.xlsx",
        "title": "Roaming Reservation RCA (Excel)",
        "type": "excel",
        "category": "OCPI Roaming"
    },
    {
        "id": "roaming_excel_all",
        "name": "ElectreeFi_Roaming_Reservation_RCA_All_Records.xlsx",
        "title": "Roaming Reservation RCA - All Records (Excel)",
        "type": "excel",
        "category": "OCPI Roaming"
    },
    {
        "id": "roaming_word",
        "name": "ElectreeFi_Roaming_Reservation_RCA_Report.docx",
        "title": "Roaming Reservation RCA Report (Word)",
        "type": "word",
        "category": "OCPI Roaming"
    },
    {
        "id": "ocpi_metrics_excel",
        "name": "ElectreeFi_OCPI_Roaming_Metrics_Analysis.xlsx",
        "title": "OCPI Roaming Metrics Analysis (Excel)",
        "type": "excel",
        "category": "OCPI Deep Dive"
    },
    {
        "id": "ocpi_metrics_30d_excel",
        "name": "ElectreeFi_OCPI_Roaming_Metrics_Analysis_30Days.xlsx",
        "title": "30-Day OCPI Roaming Metrics (Excel)",
        "type": "excel",
        "category": "OCPI Deep Dive"
    },
    {
        "id": "ocpi_deep_dive_word",
        "name": "ElectreeFi_OCPI_Roaming_Deep_Dive_Analysis.docx",
        "title": "OCPI Roaming Deep Dive Analysis (Word)",
        "type": "word",
        "category": "OCPI Deep Dive"
    },
    {
        "id": "roaming_charger_excel",
        "name": "ElectreeFi_Roaming_Charger_Wise_RCA.xlsx",
        "title": "Roaming Charger & Station RCA - IOC, MPC & ELC (Excel)",
        "type": "excel",
        "category": "OCPI Roaming"
    },
    {
        "id": "roaming_charger_word",
        "name": "ElectreeFi_Roaming_Charger_Wise_RCA_Report.docx",
        "title": "Roaming Charger & Station RCA Report - IOC, MPC & ELC (Word)",
        "type": "word",
        "category": "OCPI Roaming"
    },
    {
        "id": "uploaded_roaming_word",
        "name": "ElectreeFi_Uploaded_Roaming_RCA_Report.docx",
        "title": "Uploaded Roaming Excel - In-Detail RCA Report (Word)",
        "type": "word",
        "category": "OCPI Roaming"
    },
    {
        "id": "uploaded_roaming_excel",
        "name": "ElectreeFi_Uploaded_Roaming_RCA_Report.xlsx",
        "title": "Uploaded Roaming Excel - In-Detail Multi-Tab RCA Ledger (Excel)",
        "type": "excel",
        "category": "OCPI Roaming"
    }
]

# Task definitions corresponding to all .bat files and workflows
TASKS_CONFIG = {
    "regular_full_pipeline": {
        "title": "Run Regular RCA Full Pipeline",
        "subtitle": "Generates Cancelled Bookings Excel + 24-Hour Executive Word Report",
        "bat_equivalent": "run_regular_full_pipeline.bat",
        "category": "Regular Bookings",
        "color": "cyan",
        "steps": [
            {
                "name": "Generate Cancelled Bookings Excel",
                "cmd": [PYTHON_EXE, "generate_excel_rca_full.py"]
            },
            {
                "name": "Generate 24-Hour Executive Word Report",
                "cmd": [PYTHON_EXE, "generate_word_report.py"]
            }
        ]
    },
    "regular_excel": {
        "title": "Generate Cancelled Bookings Excel RCA",
        "subtitle": "Scrapes 24h cancellations, queries OCPP logs, classifies root causes",
        "bat_equivalent": "run_excel_rca.bat",
        "category": "Regular Bookings",
        "color": "blue",
        "steps": [
            {
                "name": "Generate Cancelled Bookings Excel",
                "cmd": [PYTHON_EXE, "generate_excel_rca_full.py"]
            }
        ]
    },
    "regular_word": {
        "title": "Generate 24-Hour Executive Word Report",
        "subtitle": "Creates formal executive RCA document (.docx) with breakdown tables",
        "bat_equivalent": "generate_word_report.bat",
        "category": "Regular Bookings",
        "color": "blue",
        "steps": [
            {
                "name": "Generate Word Report",
                "cmd": [PYTHON_EXE, "generate_word_report.py"]
            }
        ]
    },
    "quick_rca_cli": {
        "title": "Quick CLI RCA Inspection (Top 15)",
        "subtitle": "Fast live console analysis for the latest 15 cancelled bookings",
        "bat_equivalent": "run_rca.bat",
        "category": "Regular Bookings",
        "color": "slate",
        "steps": [
            {
                "name": "CLI RCA Top 15",
                "cmd": [PYTHON_EXE, "-m", "src.cli", "rca", "--max", "15"]
            }
        ]
    },
    "roaming_full_pipeline": {
        "title": "Run Roaming RCA Full Pipeline",
        "subtitle": "Generates Roaming Reservation Excel + Executive Word Report",
        "bat_equivalent": "run_roaming_full_pipeline.bat",
        "category": "OCPI Roaming",
        "color": "emerald",
        "steps": [
            {
                "name": "Generate Roaming Excel",
                "cmd": [PYTHON_EXE, "generate_roaming_rca_full.py"]
            },
            {
                "name": "Generate Roaming Word Report",
                "cmd": [PYTHON_EXE, "generate_roaming_word_report.py"]
            }
        ]
    },
    "roaming_excel": {
        "title": "Generate Roaming Reservation Excel RCA",
        "subtitle": "Scrapes OCPI Roaming bookings, fetches partner logs, extracts error codes",
        "bat_equivalent": "run_roaming_rca.bat",
        "category": "OCPI Roaming",
        "color": "emerald",
        "steps": [
            {
                "name": "Generate Roaming Excel",
                "cmd": [PYTHON_EXE, "generate_roaming_rca_full.py"]
            }
        ]
    },
    "roaming_word": {
        "title": "Generate Roaming Executive Word Report",
        "subtitle": "Generates comprehensive OCPI Roaming RCA executive document (.docx)",
        "bat_equivalent": "generate_roaming_word_report.bat",
        "category": "OCPI Roaming",
        "color": "emerald",
        "steps": [
            {
                "name": "Generate Roaming Word Report",
                "cmd": [PYTHON_EXE, "generate_roaming_word_report.py"]
            }
        ]
    },
    "ocpi_deep_excel": {
        "title": "Generate 30-Day OCPI Roaming Metrics Excel",
        "subtitle": "Comprehensive multi-tab 30-day OCPI partner breakdown & KPI matrix",
        "bat_equivalent": "generate_ocpi_excel_analysis.py",
        "category": "OCPI Deep Dive",
        "color": "indigo",
        "steps": [
            {
                "name": "Generate 30-Day OCPI Metrics Excel",
                "cmd": [PYTHON_EXE, "generate_ocpi_excel_analysis.py"]
            }
        ]
    },
    "ocpi_deep_word": {
        "title": "Generate OCPI Deep Dive Word Analysis",
        "subtitle": "In-depth executive report analyzing partner failure rates and root causes",
        "bat_equivalent": "generate_ocpi_word_analysis.py",
        "category": "OCPI Deep Dive",
        "color": "indigo",
        "steps": [
            {
                "name": "Generate OCPI Deep Dive Word Report",
                "cmd": [PYTHON_EXE, "generate_ocpi_word_analysis.py"]
            }
        ]
    },
    "roaming_charger_rca": {
        "title": "Generate Roaming Charger & Station RCA (IOC, MPC & ELC Focus)",
        "subtitle": "Calculates Charger-wise & Station-wise success/failure rates, fault attribution and hardware family analysis strictly for IOC, MPC & ELC",
        "bat_equivalent": "run_roaming_charger_rca.bat",
        "category": "OCPI Roaming",
        "color": "emerald",
        "steps": [
            {
                "name": "Generate Roaming Charger & Station RCA Excel",
                "cmd": [PYTHON_EXE, "generate_roaming_charger_rca.py"]
            }
        ]
    },
    "roaming_charger_word": {
        "title": "Generate Roaming Charger & Station Word Report (IOC, MPC & ELC Focus)",
        "subtitle": "Creates formal executive RCA document (.docx) with 4-quadrant attribution, OEM breakdown & problem charger hotlist for IOC, MPC & ELC",
        "bat_equivalent": "generate_roaming_charger_word_report.bat",
        "category": "OCPI Roaming",
        "color": "emerald",
        "steps": [
            {
                "name": "Generate Roaming Charger & Station Word Report",
                "cmd": [PYTHON_EXE, "generate_roaming_charger_word_report.py"]
            }
        ]
    },
    "safe_roaming_upload_rca": {
        "title": "Safe Clustered Roaming Upload RCA (MPC, VIN & IOC Focus)",
        "subtitle": "Filters strictly for MPC, VIN & IOC, clusters tight windows by charger, queries OCPP logs with 2s polite delay & SQLite cache, produces in-detail .docx & .xlsx (100% production safe)",
        "bat_equivalent": "run_safe_roaming_upload_rca.bat",
        "category": "OCPI Roaming",
        "color": "emerald",
        "steps": [
            {
                "name": "Run Safe Clustered Roaming Upload RCA",
                "cmd": [PYTHON_EXE, "-u", "run_safe_roaming_upload_rca.py", "--input", "uploads/latest_roaming_upload.xlsx"]
            }
        ]
    },
    "roaming_upload_rca": {
        "title": "Analyze Uploaded Roaming Excel & Generate Word + Excel RCA",
        "subtitle": "Ingests uploaded Excel (Completed & Cancelled), queries CMS Manage popup & CPO logs for IOC/VIN/MPC, generates Word (.docx) & Excel (.xlsx) reports",
        "bat_equivalent": "run_uploaded_roaming_rca.bat",
        "category": "OCPI Roaming",
        "color": "emerald",
        "steps": [
            {
                "name": "Run Uploaded Roaming Reservation RCA Analyzer",
                "cmd": [PYTHON_EXE, "-u", "run_uploaded_roaming_rca.py"]
            }
        ]
    },
    "cms_login": {
        "title": "CMS Interactive Login (CAPTCHA + OTP)",
        "subtitle": "Opens browser for operator credentials, solves CAPTCHA & saves session",
        "bat_equivalent": "login.bat",
        "category": "Authentication",
        "color": "amber",
        "steps": [
            {
                "name": "Interactive Login Window",
                "cmd": [PYTHON_EXE, "-m", "src.cli", "login", "--force"]
            }
        ]
    }
}


class TaskRunner:
    """Thread-safe runner for managing subprocesses and streaming logs."""
    def __init__(self):
        self.lock = threading.Lock()
        self.is_running = False
        self.current_task_id = None
        self.current_task_title = ""
        self.start_time = None
        self.end_time = None
        self.exit_code = 0
        self.logs = []  # list of str
        self.current_process = None
        self.should_abort = False

    def get_status(self):
        with self.lock:
            elapsed = 0
            if self.start_time:
                if self.is_running:
                    elapsed = int(time.time() - self.start_time)
                elif self.end_time:
                    elapsed = int(self.end_time - self.start_time)
            
            return {
                "running": self.is_running,
                "task_id": self.current_task_id,
                "task_title": self.current_task_title,
                "elapsed": elapsed,
                "exit_code": self.exit_code,
                "log_count": len(self.logs)
            }

    def get_logs_since(self, index: int):
        with self.lock:
            total = len(self.logs)
            if index < total:
                slice_lines = self.logs[index:]
            else:
                slice_lines = []
            
            elapsed = 0
            if self.start_time:
                if self.is_running:
                    elapsed = int(time.time() - self.start_time)
                elif self.end_time:
                    elapsed = int(self.end_time - self.start_time)

            return {
                "lines": slice_lines,
                "next_index": total,
                "running": self.is_running,
                "exit_code": self.exit_code,
                "elapsed": elapsed
            }

    def append_log(self, text: str):
        with self.lock:
            self.logs.append(text)
            if len(self.logs) > 6000:
                self.logs.pop(0)

    def clear_logs(self):
        with self.lock:
            self.logs.clear()

    def start_task(self, task_id: str):
        with self.lock:
            if self.is_running:
                return False, "Another operation is currently running."
            
            if task_id not in TASKS_CONFIG:
                return False, f"Unknown task ID: {task_id}"
            
            self.is_running = True
            self.current_task_id = task_id
            self.current_task_title = TASKS_CONFIG[task_id]["title"]
            self.start_time = time.time()
            self.end_time = None
            self.exit_code = 0
            self.should_abort = False
            self.logs.clear()
            self.logs.append(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] >>> Initializing: {self.current_task_title}")
            self.logs.append(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] >>> Working Directory: {BASE_DIR}")
            self.logs.append(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] >>> Python Binary: {PYTHON_EXE}")
            self.logs.append("-" * 75)

        thread = threading.Thread(target=self._run_task_thread, args=(task_id,), daemon=True)
        thread.start()
        return True, "Task started successfully."

    def stop_task(self):
        with self.lock:
            if not self.is_running:
                return False, "No task is currently running."
            self.should_abort = True
            if self.current_process:
                try:
                    self.current_process.terminate()
                except Exception:
                    pass
        return True, "Termination request sent."

    def _run_task_thread(self, task_id: str):
        task_cfg = TASKS_CONFIG[task_id]
        overall_exit = 0

        for idx, step in enumerate(task_cfg["steps"], start=1):
            if self.should_abort:
                self.append_log(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] [ABORTED] Operation cancelled by operator.")
                overall_exit = -1
                break

            step_name = step["name"]
            cmd = step["cmd"]
            self.append_log(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] [STEP {idx}/{len(task_cfg['steps'])}] {step_name}")
            self.append_log(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] $ {' '.join(cmd)}")

            env = os.environ.copy()
            env["PYTHONUNBUFFERED"] = "1"
            env["PYTHONIOENCODING"] = "utf-8"

            try:
                proc = subprocess.Popen(
                    cmd,
                    cwd=str(BASE_DIR),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    bufsize=1,
                    env=env
                )
                with self.lock:
                    self.current_process = proc

                for line in iter(proc.stdout.readline, ''):
                    clean_line = line.rstrip("\r\n")
                    if clean_line:
                        self.append_log(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {clean_line}")

                proc.stdout.close()
                code = proc.wait()

                if code != 0:
                    overall_exit = code
                    self.append_log(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] [ERROR] Step failed with return code {code}.")
                    break
                else:
                    self.append_log(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] [SUCCESS] Step {idx} completed successfully.")

            except Exception as e:
                self.append_log(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] [EXCEPTION] Failed to run step: {str(e)}")
                overall_exit = 1
                break

        with self.lock:
            self.is_running = False
            self.current_process = None
            self.end_time = time.time()
            self.exit_code = overall_exit
            duration = int(self.end_time - self.start_time)
            status_text = "COMPLETED SUCCESSFULLY" if overall_exit == 0 else f"FAILED (Exit Code {overall_exit})"
            self.logs.append("=" * 75)
            self.logs.append(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] >>> {status_text} in {duration}s")
            self.logs.append("=" * 75)


runner = TaskRunner()


def get_files_status():
    results = []
    for item in TRACKED_FILES:
        filepath = BASE_DIR / item["name"]
        exists = filepath.exists()
        size_str = "-"
        mod_str = "-"
        if exists:
            try:
                stat = filepath.stat()
                size_kb = stat.st_size / 1024
                if size_kb >= 1024:
                    size_str = f"{size_kb/1024:.2f} MB"
                else:
                    size_str = f"{size_kb:.1f} KB"
                mod_time = datetime.datetime.fromtimestamp(stat.st_mtime)
                mod_str = mod_time.strftime("%d-%b-%Y %H:%M:%S")
            except Exception:
                pass
        
        results.append({
            "id": item["id"],
            "name": item["name"],
            "title": item["title"],
            "type": item["type"],
            "category": item["category"],
            "exists": exists,
            "size": size_str,
            "modified": mod_str,
            "path": str(filepath)
        })
    return results


def check_session_status():
    session_file = BASE_DIR / "data" / "session_state.json"
    if not session_file.exists():
        return {
            "has_session": False,
            "status_text": "No Session File Found",
            "modified": None
        }
    try:
        stat = session_file.stat()
        mod_time = datetime.datetime.fromtimestamp(stat.st_mtime)
        return {
            "has_session": True,
            "status_text": "Session Token Stored",
            "modified": mod_time.strftime("%d-%b-%Y %H:%M")
        }
    except Exception:
        return {
            "has_session": False,
            "status_text": "Error Reading Session",
            "modified": None
        }


class ElectreeFiHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_DIR), **kwargs)

    def log_message(self, format, *args):
        # Suppress spammy HTTP request logs from console
        pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/api/status":
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            status_data = runner.get_status()
            status_data["files"] = get_files_status()
            status_data["session"] = check_session_status()
            status_data["tasks"] = TASKS_CONFIG
            status_data["python_exe"] = PYTHON_EXE
            status_data["base_dir"] = str(BASE_DIR)
            self.wfile.write(json.dumps(status_data).encode("utf-8"))
            return

        elif path == "/api/logs":
            query = urllib.parse.parse_qs(parsed.query)
            after = int(query.get("after", [0])[0])
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            log_data = runner.get_logs_since(after)
            self.wfile.write(json.dumps(log_data).encode("utf-8"))
            return

        elif path == "/api/files":
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            files_data = get_files_status()
            self.wfile.write(json.dumps(files_data).encode("utf-8"))
            return

        elif path == "/api/partner-portals":
            portals_file = BASE_DIR / "data" / "partner_portals.json"
            data = {}
            if portals_file.exists():
                try:
                    with open(portals_file, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except Exception:
                    pass
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(data).encode("utf-8"))
            return

        elif path == "/api/download-file":
            query = urllib.parse.parse_qs(parsed.query)
            fname = query.get("filename", [""])[0].strip()
            clean_name = os.path.basename(fname)
            filepath = BASE_DIR / clean_name
            if clean_name and filepath.exists() and filepath.is_file():
                self.send_response(HTTPStatus.OK)
                if clean_name.endswith(".docx"):
                    self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                elif clean_name.endswith(".xlsx"):
                    self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    self.send_header("Content-Type", "application/octet-stream")
                self.send_header("Content-Disposition", f'attachment; filename="{clean_name}"')
                self.send_header("Content-Length", str(filepath.stat().st_size))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                with open(filepath, "rb") as f:
                    while chunk := f.read(65536):
                        self.wfile.write(chunk)
                return
            else:
                self.send_error(HTTPStatus.NOT_FOUND, f"File {clean_name} not found")
                return



        # Serve UI static index.html on root
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

        if path == "/api/run":
            task_id = payload.get("task_id")
            success, message = runner.start_task(task_id)
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"success": success, "message": message}).encode("utf-8"))
            return

        elif path == "/api/stop":
            success, message = runner.stop_task()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"success": success, "message": message}).encode("utf-8"))
            return

        elif path == "/api/clear-logs":
            runner.clear_logs()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"success": True}).encode("utf-8"))
            return

        elif path == "/api/open-file":
            filename = payload.get("filename")
            filepath = BASE_DIR / filename
            success = False
            msg = ""
            if filepath.exists():
                try:
                    os.startfile(str(filepath))
                    success = True
                    msg = f"Opened {filename}"
                except Exception as e:
                    msg = str(e)
            else:
                msg = f"File {filename} does not exist yet."

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"success": success, "message": msg}).encode("utf-8"))
            return

        elif path == "/api/open-folder":
            try:
                filename = payload.get("filename")
                if filename:
                    target = BASE_DIR / filename
                    if target.exists():
                        subprocess.Popen(f'explorer /select,"{target}"')
                    else:
                        os.startfile(str(BASE_DIR))
                else:
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

        elif path == "/api/search-booking":
            if str(BASE_DIR) not in sys.path:
                sys.path.insert(0, str(BASE_DIR))
            from src.rca.booking_inspector import investigate_booking

            booking_id = str(payload.get("booking_id", "")).strip()
            start_date = str(payload.get("start_date", "")).strip()
            end_date = str(payload.get("end_date", "")).strip()

            start_date = payload.get("start_date") or None
            end_date = payload.get("end_date") or None

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

        elif path == "/api/upload-roaming-excel":
            try:
                fname = payload.get("filename", "uploaded_roaming.xlsx")
                file_type = str(payload.get("file_type", "auto")).lower().strip()
                b64_content = payload.get("content_base64", "")
                if not b64_content:
                    self.send_response(HTTPStatus.BAD_REQUEST)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": False, "message": "No file content received."}).encode("utf-8"))
                    return

                if "," in b64_content:
                    b64_content = b64_content.split(",", 1)[1]
                file_bytes = base64.b64decode(b64_content)

                uploads_dir = BASE_DIR / "uploads"
                uploads_dir.mkdir(parents=True, exist_ok=True)
                
                safe_name = re.sub(r'[^a-zA-Z0-9_.-]', '_', fname)
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                target_filename = f"{timestamp}_{safe_name}"
                saved_path = uploads_dir / target_filename
                with open(saved_path, "wb") as f:
                    f.write(file_bytes)

                # Route to dedicated slots based on file_type or filename keywords
                default_cat = None
                if file_type == "completed" or "completed" in fname.lower():
                    latest_path = uploads_dir / "latest_completed.xlsx"
                    default_cat = "Completed"
                elif file_type == "cancelled" or "cancelled" in fname.lower():
                    latest_path = uploads_dir / "latest_cancelled.xlsx"
                    default_cat = "Cancelled"
                    with open(uploads_dir / "latest_roaming_upload.xlsx", "wb") as f:
                        f.write(file_bytes)
                else:
                    latest_path = uploads_dir / "latest_roaming_upload.xlsx"

                with open(latest_path, "wb") as f:
                    f.write(file_bytes)

                if str(BASE_DIR) not in sys.path:
                    sys.path.insert(0, str(BASE_DIR))
                from src.rca.roaming_upload_analyzer import RoamingUploadAnalyzer
                analyzer = RoamingUploadAnalyzer()
                rows = analyzer.parse_uploaded_excel(str(saved_path), default_category=default_cat)

                parties = sorted(list(set(r.get("party_id", "IOC") for r in rows if r.get("party_id"))))
                completed = sum(1 for r in rows if r.get("session_category") == "Completed")
                low_kwh = sum(1 for r in rows if "complete" in r.get("session_category", "").lower() or (0.0 <= float(r.get("kwh", 0) or 0) < 1.0 and default_cat == "Completed"))
                cancelled = sum(1 for r in rows if r.get("session_category") == "Cancelled")

                summary = {
                    "total_records": len(rows),
                    "parties": parties,
                    "completed_count": completed,
                    "low_consumption_count": low_kwh,
                    "cancelled_count": cancelled,
                    "assigned_category": default_cat or "Auto-detected",
                    "preview": rows[:8]
                }

                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({
                    "success": True,
                    "filename": fname,
                    "file_type": default_cat or file_type,
                    "saved_as": target_filename,
                    "summary": summary
                }, default=str).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(HTTPStatus.INTERNAL_SERVER_ERROR)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"success": False, "message": f"Upload processing error: {str(e)}"}).encode("utf-8"))
                return

        elif path == "/api/partner-portals":
            try:
                portals_data = payload.get("portals", {})
                portals_file = BASE_DIR / "data" / "partner_portals.json"
                portals_file.parent.mkdir(parents=True, exist_ok=True)
                with open(portals_file, "w", encoding="utf-8") as f:
                    json.dump(portals_data, f, indent=2)
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"success": True, "message": "Partner portals configuration saved successfully."}).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(HTTPStatus.INTERNAL_SERVER_ERROR)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"success": False, "message": str(e)}).encode("utf-8"))
                return

        elif path == "/api/run-uploaded-roaming-rca":
            success, message = runner.start_task("roaming_upload_rca")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"success": success, "message": message}).encode("utf-8"))
            return

        elif path == "/api/login-partner":
            party_id = payload.get("party_id", "IOC").strip().upper()
            bat_file = BASE_DIR / f"login_{party_id}.bat"
            try:
                if bat_file.exists():
                    os.startfile(str(bat_file))
                else:
                    cmd = [PYTHON_EXE, "-m", "src.auth.session_manager", "--party", party_id, "--force"]
                    creation_flags = subprocess.CREATE_NEW_CONSOLE if sys.platform == "win32" else 0
                    subprocess.Popen(cmd, cwd=str(BASE_DIR), creationflags=creation_flags)
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"success": True, "message": f"Launched login window for {party_id} CMS."}).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(HTTPStatus.INTERNAL_SERVER_ERROR)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"success": False, "message": str(e)}).encode("utf-8"))
                return

        self.send_error(HTTPStatus.NOT_FOUND, "API endpoint not found")


def start_server(port: int = 58210):
    HTTPServer.allow_reuse_address = True
    server = HTTPServer(("127.0.0.1", port), ElectreeFiHandler)
    print(f"ElectreeFi Studio Server started on http://127.0.0.1:{port}")
    return server


if __name__ == "__main__":
    port = 58210
    server = start_server(port)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()


