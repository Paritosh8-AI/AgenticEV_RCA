@echo off
cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo Running ElectreeFi Root Cause Analysis (RCA)...
%PY% -m src.cli rca --max 15
pause
