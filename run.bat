@echo off
REM Text Analyzer Pro - Batch Runner
REM This script launches the GUI application with automatic setup

cd /d "%~dp0"

REM Display ASCII Banner
REM Display ASCII Banner (read from file to avoid CMD parsing issues)
cls
echo.
type "%~dp0ascii_banner.txt"
echo.
echo.
echo                       

REM Check if venv exists
if not exist .venv (
    echo [*] Virtual environment not found. Creating...
    python -m venv .venv
    if errorlevel 1 (
        echo [!] Failed to create virtual environment
        pause
        exit /b 1
    )
    echo [+] Virtual environment created successfully
    echo.
    echo [*] Installing dependencies from requirements.txt...
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\pip.exe install -r requirements.txt
    if errorlevel 1 (
        echo [!] Failed to install dependencies
        pause
        exit /b 1
    )
    echo [+] Dependencies installed successfully
    echo.
) else (
    echo [+] Virtual environment found. Checking for updates...
    .venv\Scripts\pip.exe install -r requirements.txt --quiet
)

echo [*] Launching Text Analyzer Pro...
echo.
.venv\Scripts\python.exe gui.py
pause
