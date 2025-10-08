@echo off
echo ========================================
echo Hardware Information Collection System
echo ========================================
echo.

REM Step 1: Collect hardware information
echo [Step 1/2] Collecting hardware information...
echo.
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\systeminfo.ps1"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Failed to collect hardware information.
    pause
    exit /b 1
)

echo.
echo Hardware information collected successfully.
echo.

REM Step 2: Check if config.json exists
if not exist "%~dp0config.json" (
    echo [Step 2/2] Skipping Google Spreadsheet sync - config.json not found.
    echo.
    echo To enable automatic sync to Google Spreadsheet:
    echo 1. Copy config.json.example to config.json
    echo 2. Follow the setup instructions in README.md
    echo 3. Run this script again
    goto :done
)

REM Step 2: Sync to Google Spreadsheet
echo [Step 2/2] Syncing to Google Spreadsheet...
echo.
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\sync_to_spreadsheet_v2.ps1"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo WARNING: Failed to sync to Google Spreadsheet.
    echo CSV file has been created locally at: hardware_info.csv
    echo.
    echo To troubleshoot:
    echo - Check your config.json file
    echo - Verify your Web App URL is correct
    echo - See README.md for more information
    goto :done
)

echo.
echo Successfully synced to Google Spreadsheet!

:done
echo.
echo ========================================
echo All processes completed.
echo ========================================
pause
