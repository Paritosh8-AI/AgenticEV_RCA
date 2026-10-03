@echo off
cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo ============================================================
echo   ElectreeFi - Run Uploaded Roaming Reservation RCA
echo   Focus: IOC, VIN, MPC ^| 100%% Production Safe
echo ============================================================
%PY% -u run_uploaded_roaming_rca.py
pause
