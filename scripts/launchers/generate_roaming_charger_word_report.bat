@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)


echo ===============================================================================
echo ELECTREEFI CMS - ROAMING CHARGER-WISE AND STATION-WISE WORD REPORT (.docx)
echo ===============================================================================
echo.

%PY% generate_roaming_charger_word_report.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Generation failed with error code %ERRORLEVEL%.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ===============================================================================
echo SUCCESS: WORD REPORT CREATED SUCCESSFULLY!
echo Location: .\ElectreeFi_Roaming_Charger_Wise_RCA_Report.docx
echo ===============================================================================
echo.
pause
