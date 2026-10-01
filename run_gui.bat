@echo off
cd /d "%~dp0"
echo ============================================================
echo   Launching ElectreeFi Automation & RCA Studio...
echo ============================================================
if exist ".\ElectreeFi_Hub.exe" (
    start "" ".\ElectreeFi_Hub.exe"
    exit /b 0
)
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)
%PY% electreefi_app.py
