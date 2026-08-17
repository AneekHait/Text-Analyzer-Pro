@echo off
REM Text Analyzer Pro - Launcher
REM First run: creates venv and installs dependencies
REM Subsequent runs: launches the app directly

cd /d "%~dp0"
cls
echo.
type "%~dp0ascii_banner.txt"
echo.
echo.

REM Check if venv exists
if not exist .venv\Scripts\python.exe (
    echo ============================================================
    echo   FIRST RUN - Setting up environment...
    echo ============================================================
    echo.
    echo [1/3] Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo [!] ERROR: Failed to create virtual environment.
        echo     Make sure Python 3.10+ is installed and on PATH.
        pause
        exit /b 1
    )
    echo       Done.
    echo.
    echo [2/3] Upgrading pip...
    .venv\Scripts\python.exe -m pip install --upgrade pip --quiet
    echo       Done.
    echo.
    echo [3/3] Installing dependencies (this may take a few minutes)...
    .venv\Scripts\pip.exe install -r requirements.txt
    if errorlevel 1 (
        echo [!] ERROR: Failed to install dependencies.
        pause
        exit /b 1
    )
    echo.
    echo ============================================================
    echo   Setup complete! Launching Text Analyzer Pro...
    echo ============================================================
    echo.
) else (
    echo [*] Launching Text Analyzer Pro...
    echo.
)

.venv\Scripts\python.exe gui.py
if errorlevel 1 (
    echo.
    echo [!] Application exited with an error.
    pause
)
