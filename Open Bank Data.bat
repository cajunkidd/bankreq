@echo off
REM Stine Bank Data Viewer launcher.
REM Bootstraps a local Python venv on first run, then starts the Streamlit
REM web app and opens it in the default browser.
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "APP=app.py"
set "REQS=requirements.txt"
set "VENV=.venv"
set "PYEXE=%VENV%\Scripts\python.exe"
set "PORT=8501"

if not exist "%APP%" (
    echo [ERROR] Could not find "%APP%" in this folder:
    echo   %~dp0
    echo Make sure all program files stayed together.
    pause
    exit /b 1
)

REM --- Locate a Python launcher ---------------------------------------------
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo [ERROR] Python is not installed or not on PATH.
    echo.
    echo Please install Python 3.10 or newer from:
    echo     https://www.python.org/downloads/
    echo During install, check "Add Python to PATH", then run this file again.
    pause
    exit /b 1
)

REM --- Create venv on first run ---------------------------------------------
if not exist "%PYEXE%" (
    echo First-time setup: creating local Python environment...
    %PY% -m venv "%VENV%"
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
)

REM --- Install / update dependencies ----------------------------------------
if not exist "%VENV%\.deps_installed" (
    echo Installing dependencies ^(one-time, ~30-60 seconds^)...
    "%PYEXE%" -m pip install --upgrade pip >nul
    "%PYEXE%" -m pip install -r "%REQS%"
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies.
        pause
        exit /b 1
    )
    echo. > "%VENV%\.deps_installed"
)

REM --- Launch Streamlit and open browser ------------------------------------
echo Starting Bank Data Viewer at http://localhost:%PORT% ...
echo Close this window to stop the app.
echo.
start "" "http://localhost:%PORT%"
"%PYEXE%" -m streamlit run "%APP%" --server.port %PORT% --server.headless true --browser.gatherUsageStats false

endlocal
