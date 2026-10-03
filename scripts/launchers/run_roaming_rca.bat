@echo off
cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo ============================================================
echo ElectreeFi - Roaming Reservation RCA Excel Generator
echo ============================================================
echo.
%PY% generate_roaming_rca_full.py
echo.
pause
