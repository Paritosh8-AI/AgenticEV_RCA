@echo off
title ElectreeFi - IndianOil e-Charge (IOC) Login
cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo ============================================================
echo   IndianOil e-Charge (IOC) - CMS Authentication Portal
echo ============================================================
echo.
echo URL: https://cpo-ioc.ev-network.com/Account/Login
echo Username: cpo_ioc_admin
echo.
echo Launching visible browser window...
echo Please solve the visual CAPTCHA in the browser and click Login.
echo Session cookies will automatically be saved to data\session_IOC.json.
echo.
%PY% -u -m src.auth.session_manager --party IOC --force
echo.
pause
