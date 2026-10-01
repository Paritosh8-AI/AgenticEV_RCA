@echo off
cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo ============================================================
echo ElectreeFi - Cancelled Bookings RCA Excel Generator (24h)
echo ============================================================
echo.
%PY% generate_excel_rca_full.py
echo.
pause
