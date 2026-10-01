@echo off
cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo ============================================================
echo Starting ElectreeFi Login Window...
echo Solve the CAPTCHA and enter your email OTP in the browser.
echo ============================================================
%PY% -m src.cli login --force
echo.
echo Login process finished.
pause
