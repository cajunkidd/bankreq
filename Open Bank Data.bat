@echo off
REM Stine BankReq Reformatter - self-bootstrapping launcher.
REM Drop this .bat file into ANY folder and double-click it. It will:
REM   1. Auto-install Python 3.13 (per-user, no admin) if no supported
REM      Python is detected.
REM   2. Download the latest app.py, requirements.txt, and logo from GitHub.
REM   3. Create a local Python virtual environment on first run.
REM   4. Install dependencies (one-time).
REM   5. Launch the app in your default browser at http://localhost:8501.

setlocal EnableExtensions EnableDelayedExpansion

echo.
echo === Stine BankReq Reformatter launcher ===
echo.

REM cmd.exe refuses to use a UNC path as the current directory, which is
REM common when the Desktop is on a roaming network share. pushd maps
REM UNC paths to a temporary drive letter so the rest of the script can run.
pushd "%~dp0" 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] Could not enter the folder containing this script:
    echo   %~dp0
    echo.
    echo If that path starts with two backslashes it is a network share.
    echo Try copying "Open Bank Data.bat" to a local folder like C:\BankReq
    echo and running it from there instead.
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
set "PY_VERSION=3.13.3"
set "PY_URL=https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%-amd64.exe"

REM Disable pip cache. AppData is sometimes on a network share and the
REM cache then triggers "Permission denied" errors.
set "PIP_NO_CACHE_DIR=1"
set "PIP_DISABLE_PIP_VERSION_CHECK=1"

REM ---- Locate a supported Python interpreter -----------------------------
set "PY="
call :find_py 3.13
if not defined PY call :find_py 3.12
if not defined PY call :find_py 3.11
if not defined PY call :install_python
if not defined PY (
    echo.
    echo [ERROR] Could not install or locate Python 3.13.
    echo.
    echo Please install Python 3.13 manually from:
    echo   https://www.python.org/downloads/
    echo and check "Add Python to PATH" during the installer.
    echo.
    pause
    exit /b 1
)
echo Using Python at: !PY!
echo.

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
    echo [WARNING] Download failed - using previously cached %APP%.
)
call :fetch "%REQS%" "%BASE_URL%/%REQS%"
call :fetch "%LOGO%" "%BASE_URL%/%LOGO%"

REM ---- Create venv + install deps on first run ----------------------------
REM If a previous venv used an incompatible Python (e.g. 3.15), recreate it.
if exist "%PYEXE%" (
    "%PYEXE%" -c "import sys; ok = sys.version_info[0]==3 and 11 <= sys.version_info[1] <= 13; raise SystemExit(0 if ok else 1)" 2>nul
    if errorlevel 1 (
        echo Existing virtualenv uses an incompatible Python version - recreating...
        rmdir /s /q "%VENV%"
    )
)

if not exist "%PYEXE%" (
    echo First-time setup: creating local Python environment...
    "!PY!" -m venv "%VENV%" || (
        echo [ERROR] venv creation failed.
        pause
        exit /b 1
    )
)

if not exist "%VENV%\.deps_installed" (
    echo Installing dependencies, one-time, about 30 to 60 seconds...
    "%PYEXE%" -m pip install --no-cache-dir --upgrade pip >nul
    "%PYEXE%" -m pip install --no-cache-dir -r "%REQS%" || (
        echo [ERROR] pip install failed.
        pause
        exit /b 1
    )
    echo. > "%VENV%\.deps_installed"
)

REM ---- Launch -------------------------------------------------------------
REM Pre-create streamlit credentials so the first-run email prompt does not
REM block startup. Without this, streamlit asks for an email on stdin the
REM very first time it runs, which prevents the browser from auto-opening.
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit" >nul 2>&1
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    > "%USERPROFILE%\.streamlit\credentials.toml" echo [general]
    >> "%USERPROFILE%\.streamlit\credentials.toml" echo email = ""
)

echo.
echo Starting Stine BankReq Reformatter at http://localhost:%PORT% ...
echo Your browser will open automatically once the server is ready.
echo Close this window to stop the app.
echo.

REM Drop --server.headless so streamlit polls its own readiness and opens
REM the default browser via Python's webbrowser module. This is the most
REM reliable autolaunch path - no PowerShell race, no execution-policy gotchas.
"%PYEXE%" -m streamlit run "%APP%" --server.port %PORT% --server.address localhost --browser.serverAddress localhost --browser.gatherUsageStats false
set "STREAMLIT_EXIT=%errorlevel%"

echo.
echo ============================================================
echo Streamlit exited with code %STREAMLIT_EXIT%.
echo If the app stopped unexpectedly, the error message is above.
echo ============================================================
echo.
pause

popd
endlocal
exit /b %STREAMLIT_EXIT%


REM =========================================================================
REM :find_py <version>
REM Looks up python.exe for the requested major.minor and stores the full
REM path in PY. Does nothing if the interpreter is not available.
REM =========================================================================
:find_py
for /f "delims=" %%P in ('py -%1 -c "import sys; print(sys.executable)" 2^>nul') do set "PY=%%P"
exit /b 0

REM =========================================================================
REM :install_python
REM Downloads the official Python installer and runs it per-user (no admin
REM required). On success, stores the resulting python.exe path in PY.
REM =========================================================================
:install_python
echo No supported Python version was found on this computer.
echo Downloading Python %PY_VERSION% installer, about 28 MB...
set "PY_INSTALLER=%TEMP%\python-%PY_VERSION%-installer.exe"
call :fetch "%PY_INSTALLER%" "%PY_URL%"
if errorlevel 1 (
    echo [ERROR] Could not download the Python installer.
    echo Check your internet connection and try again.
    exit /b 0
)
echo Installing Python %PY_VERSION% per-user. No admin password is needed.
echo A small progress bar may appear briefly.
"%PY_INSTALLER%" /passive InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1
del "%PY_INSTALLER%" 2>nul
REM Locate the freshly installed Python. Per-user installs land in
REM %LOCALAPPDATA%\Programs\Python\Python313\.
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
)
exit /b 0

REM =========================================================================
REM :fetch <local-path> <remote-url>
REM Downloads the remote URL into the local path. Tries curl first, then
REM PowerShell. Returns errorlevel 0 on success, non-zero on failure.
REM =========================================================================
:fetch
where curl >nul 2>nul
if not errorlevel 1 (
    curl -fsSL -o %1 %2
    exit /b %errorlevel%
)
powershell -NoProfile -Command "try { Invoke-WebRequest -Uri %2 -OutFile %1 -UseBasicParsing; exit 0 } catch { exit 1 }"
exit /b %errorlevel%
