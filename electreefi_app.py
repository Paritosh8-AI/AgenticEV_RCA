"""
ElectreeFi Automation & RCA Studio
Main Desktop Executable Entry Point
"""

import os
import sys
import time
import socket
import pathlib
import threading
import subprocess
import webbrowser

if getattr(sys, "frozen", False):
    BASE_DIR = pathlib.Path(sys.executable).resolve().parent
else:
    BASE_DIR = pathlib.Path(__file__).resolve().parent

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from gui.server import start_server


def find_free_port(start_port: int = 58210) -> int:
    for port in range(start_port, start_port + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    return start_port


def find_edge_or_chrome():
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe")
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def launch_window(url: str):
    browser_exe = find_edge_or_chrome()
    profile_dir = BASE_DIR / "data" / ".app_profile"
    profile_dir.mkdir(parents=True, exist_ok=True)

    if browser_exe:
        cmd = [
            browser_exe,
            f"--app={url}",
            "--window-size=1380,880",
            f"--user-data-dir={profile_dir}"
        ]
        try:
            return subprocess.Popen(cmd)
        except Exception as e:
            print(f"Could not launch browser app mode: {e}, falling back to default browser.")

    webbrowser.open(url)
    return None


def main():
    port = find_free_port(58210)
    server = start_server(port)
    url = f"http://127.0.0.1:{port}"
    print(f"\n=======================================================")
    print(f"  ElectreeFi Automation & RCA Studio")
    print(f"  Interface running at: {url}")
    print(f"=======================================================\n")

    # Launch desktop window
    launch_window(url)

    # Keep server alive in main thread
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down ElectreeFi Studio...")
    finally:
        server.shutdown()


if __name__ == "__main__":
    main()
