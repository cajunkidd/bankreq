@echo off
REM Launcher: opens the bank data spreadsheet in the user's default app.
setlocal
cd /d "%~dp0"

set "FILE=raw data (2).xlsx"

if not exist "%FILE%" (
    echo Could not find "%FILE%" in:
    echo   %~dp0
    echo.
    echo Make sure this .bat file is in the same folder as the spreadsheet.
    pause
    exit /b 1
)

start "" "%FILE%"
endlocal
