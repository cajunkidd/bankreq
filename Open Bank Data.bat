@echo off
REM Developer / "run from source" launcher.
REM End users do NOT need this file — they get BankDataViewer.exe from the
REM GitHub Releases page and double-click it.
REM
REM This .bat exists for local development: it creates a venv, installs
REM dependencies, and runs app.py directly.
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "APP=app.py"
set "REQS=requirements.txt"
set "VENV=.venv"
set "PYEXE=%VENV%\Scripts\pythonw.exe"

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
    echo Installing dependencies...
    "%VENV%\Scripts\python.exe" -m pip install --upgrade pip >nul
    "%VENV%\Scripts\python.exe" -m pip install -r "%REQS%" || (echo [ERROR] pip install failed & pause & exit /b 1)
    echo. > "%VENV%\.deps_installed"
)

start "" "%PYEXE%" "%APP%"
endlocal
