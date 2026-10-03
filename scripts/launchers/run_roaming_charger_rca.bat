@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)


echo ===============================================================================
echo ELECTREEFI CMS - ROAMING CHARGER-WISE AND STATION-WISE RCA REPORT
echo (Success and Failure Rates, Fault Attribution: Charger vs User vs CMS)
echo ===============================================================================
echo.

%PY% generate_roaming_charger_rca.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Generation failed with error code %ERRORLEVEL%.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ===============================================================================
echo SUCCESS: REPORT CREATED SUCCESSFULLY!
echo Location: .\ElectreeFi_Roaming_Charger_Wise_RCA.xlsx
echo ===============================================================================
echo.
pause
