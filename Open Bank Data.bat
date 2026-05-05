@echo off
REM Developer / "run from source" launcher.
REM End users do NOT need this file — they get BankDataViewer.exe from the
REM GitHub Releases page and double-click it.
REM
REM This .bat exists for local development: it creates a venv, installs
REM dependencies, and runs the Streamlit app directly in the browser.
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "APP=app.py"
set "REQS=requirements.txt"
set "VENV=.venv"
set "PYEXE=%VENV%\Scripts\python.exe"
set "PORT=8501"

if not exist "%APP%" (
    echo [ERROR] Could not find "%APP%" in this folder.
    pause
    exit /b 1
)

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo [ERROR] Python is not installed or not on PATH.
    echo Install Python 3.10+ from https://www.python.org/downloads/
    pause
    exit /b 1
)

if not exist "%PYEXE%" (
    echo First-time setup: creating local Python environment...
    %PY% -m venv "%VENV%" || (echo [ERROR] venv creation failed & pause & exit /b 1)
)

if not exist "%VENV%\.deps_installed" (
    echo Installing dependencies ^(one-time, ~30-60 seconds^)...
    "%PYEXE%" -m pip install --upgrade pip >nul
    "%PYEXE%" -m pip install -r "%REQS%" || (echo [ERROR] pip install failed & pause & exit /b 1)
    echo. > "%VENV%\.deps_installed"
)

echo Starting Bank Data Viewer at http://localhost:%PORT% ...
echo Close this window to stop the app.
echo.
start "" "http://localhost:%PORT%"
"%PYEXE%" -m streamlit run "%APP%" --server.port %PORT% --server.headless true --browser.gatherUsageStats false

endlocal
