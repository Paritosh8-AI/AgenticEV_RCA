@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)


echo ===============================================================================
echo ELECTREEFI CMS - REGULAR BOOKINGS RCA COMPLETE PIPELINE
echo (Generates both Cancelled Bookings Excel Spreadsheet and Executive Word Report)
echo ===============================================================================
echo.

echo [OPERATION 1/2] Fetching Cancelled Bookings and Generating Excel Spreadsheet...
echo -------------------------------------------------------------------------------
%PY% generate_excel_rca_full.py
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
%PY% generate_word_report.py
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
echo 1. Excel Workbook : .\ElectreeFi_Cancelled_Bookings_RCA.xlsx
echo 2. Word Report    : .\ElectreeFi_24Hour_RCA_Report.docx
echo ===============================================================================
echo.
pause
