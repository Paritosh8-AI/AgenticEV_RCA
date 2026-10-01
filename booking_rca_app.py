"""
ElectreeFi Booking RCA Deep Dive Studio
Dedicated Standalone Desktop Application Entry Point
"""

import os
import sys
import socket
import pathlib
import subprocess
import webbrowser

if getattr(sys, "frozen", False):
    exe_dir = pathlib.Path(sys.executable).resolve().parent
    if (exe_dir / "data").exists():
        BASE_DIR = exe_dir
    elif (exe_dir.parent / "data").exists():
        BASE_DIR = exe_dir.parent
    else:
        BASE_DIR = exe_dir
else:
    BASE_DIR = pathlib.Path(__file__).resolve().parent

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from gui.booking_rca_server import start_server


def find_free_port(start_port: int = 58220) -> int:
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
    profile_dir = BASE_DIR / "data" / ".booking_rca_profile"
    profile_dir.mkdir(parents=True, exist_ok=True)

    if browser_exe:
        cmd = [
            browser_exe,
            f"--app={url}",
            "--window-size=1420,900",
            f"--user-data-dir={profile_dir}"
        ]
        try:
            return subprocess.Popen(cmd)
        except Exception as e:
            print(f"Could not launch browser app mode: {e}, falling back to default browser.")

    webbrowser.open(url)
    return None


def main():
    port = find_free_port(58220)
    server = start_server(port)
    url = f"http://127.0.0.1:{port}"
    print(f"\n=======================================================")
    print(f"  ElectreeFi Booking RCA Deep Dive Studio (Standalone)")
    print(f"  Interface running at: {url}")
    print(f"=======================================================\n")

    # Launch desktop window
    launch_window(url)

    # Keep server alive in main thread
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down Booking RCA Studio...")
    finally:
        server.shutdown()


if __name__ == "__main__":
    main()
