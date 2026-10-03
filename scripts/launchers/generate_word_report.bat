@echo off
cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo Generating ElectreeFi 24-Hour Hourly RCA Word Report (.docx)...
%PY% -m src.cli export-report --output ElectreeFi_24Hour_RCA_Report.docx
pause
