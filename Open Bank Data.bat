@echo off
REM Stine BankReq Reformatter — self-bootstrapping launcher.
REM Drop this .bat file into ANY folder and double-click it. It will:
REM   1. Verify Python 3 is installed (and prompt to install if not).
REM   2. Download the latest app.py, requirements.txt, and logo from GitHub.
REM   3. Create a local Python virtual environment on first run.
REM   4. Install dependencies (one-time).
REM   5. Launch the app in your default browser at http://localhost:8501.
REM
REM Requires Python 3.10+ from https://www.python.org/downloads/
REM (be sure to check "Add Python to PATH" during install).

setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "APP=app.py"
set "REQS=requirements.txt"
set "LOGO=Stinelogo_white_rec.svg"
set "VENV=.venv"
set "PYEXE=%VENV%\Scripts\python.exe"
set "PORT=8501"
set "BRANCH=claude/add-features-improvements-f9al2"
set "BASE_URL=https://raw.githubusercontent.com/cajunkidd/bankreq/%BRANCH%"

REM ---- Verify Python ------------------------------------------------------
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo.
    echo [ERROR] Python is not installed or not on PATH.
    echo.
    echo Install Python 3.10+ from https://www.python.org/downloads/
    echo IMPORTANT: check "Add Python to PATH" during the installer.
    echo.
    pause
    exit /b 1
)

REM ---- Fetch latest source files from GitHub ------------------------------
echo Checking for the latest version on GitHub...
call :fetch "%APP%"  "%BASE_URL%/%APP%"
if errorlevel 1 (
    if not exist "%APP%" (
        echo.
        echo [ERROR] Could not download %APP% and no cached copy exists.
        echo Check your internet connection and try again.
        echo.
        pause
        exit /b 1
    )
    echo [WARNING] Download failed — using previously cached %APP%.
)
call :fetch "%REQS%" "%BASE_URL%/%REQS%"
call :fetch "%LOGO%" "%BASE_URL%/%LOGO%"

REM ---- Create venv + install deps on first run ----------------------------
if not exist "%PYEXE%" (
    echo First-time setup: creating local Python environment...
    %PY% -m venv "%VENV%" || (
        echo [ERROR] venv creation failed.
        pause
        exit /b 1
    )
)

if not exist "%VENV%\.deps_installed" (
    echo Installing dependencies ^(one-time, ~30-60 seconds^)...
    "%PYEXE%" -m pip install --upgrade pip >nul
    "%PYEXE%" -m pip install -r "%REQS%" || (
        echo [ERROR] pip install failed.
        pause
        exit /b 1
    )
    echo. > "%VENV%\.deps_installed"
)

REM ---- Launch -------------------------------------------------------------
echo.
echo Starting Stine BankReq Reformatter at http://localhost:%PORT% ...
echo Close this window to stop the app.
echo.
start "" "http://localhost:%PORT%"
"%PYEXE%" -m streamlit run "%APP%" --server.port %PORT% --server.headless true --browser.gatherUsageStats false

endlocal
exit /b 0


REM =========================================================================
REM :fetch <local-path> <remote-url>
REM Downloads the remote URL into the local path. Tries curl first (built into
REM Windows 10 1803+ and Windows 11), falls back to PowerShell.
REM Returns errorlevel 0 on success, non-zero on failure.
REM =========================================================================
:fetch
where curl >nul 2>nul
if not errorlevel 1 (
    curl -fsSL -o %1 %2
    exit /b %errorlevel%
)
powershell -NoProfile -Command "try { Invoke-WebRequest -Uri %2 -OutFile %1 -UseBasicParsing; exit 0 } catch { exit 1 }"
exit /b %errorlevel%
