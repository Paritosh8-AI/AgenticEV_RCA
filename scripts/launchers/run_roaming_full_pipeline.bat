@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)


echo ===============================================================================
echo ELECTREEFI CMS - ROAMING RCA COMPLETE AUTOMATION PIPELINE
echo (Generates both Roaming Excel Workbook and Executive Word Report)
echo ===============================================================================
echo.

echo [OPERATION 1/2] Fetching Roaming Bookings and Generating Excel Spreadsheet...
echo -------------------------------------------------------------------------------
%PY% generate_roaming_rca_full.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Operation 1 (Excel Generation) failed with error code %ERRORLEVEL%.
    echo Aborting pipeline.
    echo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [OPERATION 2/2] Generating Executive Word Report (.docx)...
echo -------------------------------------------------------------------------------
%PY% generate_roaming_word_report.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Operation 2 (Word Report Generation) failed with error code %ERRORLEVEL%.
    echo Aborting pipeline.
    echo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ===============================================================================
echo SUCCESS: BOTH OPERATIONS COMPLETED SUCCESSFULLY!
echo ===============================================================================
echo 1. Excel Workbook : .\ElectreeFi_Roaming_Reservation_RCA.xlsx
echo 2. Word Report    : .\ElectreeFi_Roaming_Reservation_RCA_Report.docx
echo ===============================================================================
echo.
pause
