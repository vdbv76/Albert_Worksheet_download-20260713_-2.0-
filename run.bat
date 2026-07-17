@echo off
REM ============================================================================
REM  Albert Worksheet Download - Streamlit launcher
REM  Double-click this file to start the app in your browser.
REM ============================================================================

setlocal

REM --- Always run from the folder this .bat lives in --------------------------
cd /d "%~dp0"

REM --- Find a Python launcher -------------------------------------------------
REM  1) A known local Python 3.12 install (this machine keeps it here, off PATH).
REM     Edit PYHOME below if your Python lives somewhere else.
REM  2) Otherwise the "py" launcher with -3 (newest installed Python 3.x).
REM  3) Otherwise plain "python" on PATH.
set "PY="
set "PYHOME=C:\Users\bujvilal\python312"
if exist "%PYHOME%\python.exe" set "PY="%PYHOME%\python.exe""
if not defined PY (
    where py >nul 2>nul && set "PY=py -3"
)
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo [ERROR] Python was not found on your PATH.
    echo         Install Python 3.10 or newer from https://www.python.org/downloads/
    echo         and be sure to tick "Add Python to PATH" during setup.
    echo.
    pause
    exit /b 1
)

echo Using Python launcher: %PY%
echo.

REM --- Require a modern Python (3.10+ for streamlit + truststore) -------------
REM  Old Pythons (e.g. 3.5) fail here with a confusing "SSL: CERTIFICATE_
REM  VERIFY_FAILED" during pip install, because their pip points at the retired
REM  pypi.python.org and their bundled certificates are too old to validate it.
REM  Catch that up front and tell the user what to actually do.
%PY% -c "import sys; sys.exit(0 if sys.version_info[:2] >= (3,10) else 1)" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Your Python is too old for this app.
    %PY% -c "import platform; print('        Detected Python ' + platform.python_version() + '.')"
    echo         This app needs Python 3.10 or newer ^(streamlit and truststore
    echo         do not support older versions^).
    echo.
    echo         Please install Python 3.11 or 3.12 from
    echo             https://www.python.org/downloads/
    echo         tick "Add Python to PATH" during setup, then run this file again.
    echo.
    pause
    exit /b 1
)

REM --- Make sure ALL requirements are installed ------------------------------
REM  Check every dependency, not just streamlit: this machine already had
REM  streamlit but was missing truststore, so a streamlit-only check would
REM  wrongly skip installation. find_spec avoids actually importing (fast, no
REM  side effects) and returns non-zero if ANY package is missing.
%PY% -c "import importlib.util as u, sys; sys.exit(1 if any(u.find_spec(m) is None for m in ['streamlit','albert','pandas','truststore','openpyxl']) else 0)" >nul 2>nul
if errorlevel 1 (
    echo Some dependencies are missing - installing from requirements.txt ...
    %PY% -m pip install --upgrade pip
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [ERROR] Dependency installation failed. See the messages above.
        pause
        exit /b 1
    )
)

REM --- Launch the app --------------------------------------------------------
echo.
echo Starting the Albert Worksheet app...
echo A browser tab should open automatically. Close this window to stop the app.
echo.
%PY% -m streamlit run app.py

REM --- If Streamlit exits with an error, keep the window open -----------------
if errorlevel 1 (
    echo.
    echo [ERROR] Streamlit exited unexpectedly. See the messages above.
    pause
)

endlocal
