@echo off
REM Stine Bank Reconciliation - Windows launcher (run from source)
REM Requires Python 3.10+ on PATH.

setlocal
cd /d "%~dp0"

if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 goto :err
)

call .venv\Scripts\activate.bat
if errorlevel 1 goto :err

echo Installing/updating dependencies...
python -m pip install --quiet -r requirements.txt
if errorlevel 1 goto :err

echo Starting the app...
python run.py
goto :eof

:err
echo.
echo Something went wrong. Make sure Python 3.10+ is installed
echo and on your PATH (https://www.python.org/downloads/).
pause
