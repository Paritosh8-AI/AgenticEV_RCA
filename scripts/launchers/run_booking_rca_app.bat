@echo off
setlocal enabledelayedexpansion
title ElectreeFi Booking RCA Deep Dive Studio

cd /d "%~dp0"
if exist ".\.venv\Scripts\python.exe" (
    set "PY=.\.venv\Scripts\python.exe"
) else (
    set "PY=python"
)


echo ===============================================================================
echo ELECTREEFI - BOOKING RCA DEEP DIVE STUDIO (STANDALONE)
echo ===============================================================================
echo.

if exist ".\ElectreeFi_Booking_RCA.exe" (
    echo Launching standalone binary ElectreeFi_Booking_RCA.exe...
    start "" ".\ElectreeFi_Booking_RCA.exe"
    exit /b 0
)

if exist ".\dist\ElectreeFi_Booking_RCA.exe" (
    echo Launching standalone binary dist\ElectreeFi_Booking_RCA.exe...
    start "" ".\dist\ElectreeFi_Booking_RCA.exe"
    exit /b 0
)

if exist "%PY%" (
    set "PYTHON_CMD=%PY%"
) else (
    set "PYTHON_CMD=python"
)

echo Launching Booking RCA Deep Dive Studio via Python...
%PYTHON_CMD% booking_rca_app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Application exited with error code %ERRORLEVEL%.
    pause
)
