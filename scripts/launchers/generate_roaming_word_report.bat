@echo off
cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo ============================================================
echo ElectreeFi - Roaming Reservation RCA Word Report (.docx)
echo ============================================================
echo.
%PY% generate_roaming_word_report.py
echo.
pause
