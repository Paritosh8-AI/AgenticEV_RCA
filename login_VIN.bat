@echo off
title ElectreeFi - VinFast Auto (VIN) Login
cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo ============================================================
echo   VinFast Auto (VIN) - CMS Authentication Portal
echo ============================================================
echo.
echo URL: https://cpo-vin.ev-network.com/Account/Login
echo Username: cpo_operator
echo.
echo Launching visible browser window...
echo Please solve the visual CAPTCHA in the browser and click Login.
echo Session cookies will automatically be saved to data\session_VIN.json.
echo.
%PY% -u -m src.auth.session_manager --party VIN --force
echo.
pause
