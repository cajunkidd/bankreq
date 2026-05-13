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

REM cmd.exe refuses to use a UNC path (\\server\share\...) as the current
REM directory, which is common when a user keeps this .bat on a roaming or
REM redirected Desktop. pushd transparently maps a UNC path to a temporary
REM drive letter so the rest of the script can run.
pushd "%~dp0" 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] Could not enter the folder containing this script:
    echo   %~dp0
    echo.
    echo If that path starts with "\\" it is a network share. Try copying
    echo "Open Bank Data.bat" to a local folder like C:\BankReq and running
    echo it from there instead.
    echo.
    pause
    exit /b 1
)

set "APP=app.py"
set "REQS=requirements.txt"
set "LOGO=Stinelogo_white_rec.svg"
set "VENV=.venv"
set "PYEXE=%VENV%\Scripts\python.exe"
set "PORT=8501"
set "BRANCH=claude/add-features-improvements-f9al2"
set "BASE_URL=https://raw.githubusercontent.com/cajunkidd/bankreq/%BRANCH%"

REM ---- Verify Python ------------------------------------------------------
REM Streamlit's transitive deps (pyarrow, pandas) ship prebuilt wheels for
REM Python 3.11 / 3.12 / 3.13. Versions like 3.14 / 3.15 force a source
REM build that needs a C++ compiler and almost always fails on user
REM machines, so we explicitly prefer the supported versions.
set "PY="
py -3.13 -c "" >nul 2>nul && set "PY=py -3.13"
if not defined PY (
    py -3.12 -c "" >nul 2>nul && set "PY=py -3.12"
)
if not defined PY (
    py -3.11 -c "" >nul 2>nul && set "PY=py -3.11"
)
if not defined PY (
    echo.
    echo [ERROR] Python 3.11, 3.12, or 3.13 is required.
    echo.
    echo Install Python 3.13 from https://www.python.org/downloads/
    echo IMPORTANT: check "Add Python to PATH" during the installer.
    echo.
    echo (Newer Python versions like 3.14 or 3.15 are not yet supported by
    echo  all required packages and will fail with a build error.)
    echo.
    pause
    exit /b 1
)
echo Using %PY%.

REM Some users have AppData redirected to a network share, which causes
REM "Permission denied" errors when pip tries to read or write its wheel
REM cache. Disable the cache entirely so pip never touches that folder.
set "PIP_NO_CACHE_DIR=1"
set "PIP_DISABLE_PIP_VERSION_CHECK=1"

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
REM If a venv from a previous run uses an incompatible Python (e.g. 3.15),
REM nuke it so we can rebuild against the supported interpreter found above.
if exist "%PYEXE%" (
    "%PYEXE%" -c "import sys; raise SystemExit(0 if (3,11) <= sys.version_info[:2] <= (3,13) else 1)" 2>nul
    if errorlevel 1 (
        echo Existing virtualenv uses an incompatible Python version - recreating...
        rmdir /s /q "%VENV%"
    )
)

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
    "%PYEXE%" -m pip install --no-cache-dir --upgrade pip >nul
    "%PYEXE%" -m pip install --no-cache-dir -r "%REQS%" || (
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

popd
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
