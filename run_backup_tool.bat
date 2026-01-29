@echo off
echo ========================================
echo Outlook Email Backup Tool
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7 or higher from python.org
    pause
    exit /b 1
)

echo Python found. Starting application...
echo.

REM Run the application
python main.py

REM Pause if there's an error
if errorlevel 1 (
    echo.
    echo Application exited with an error.
    pause
)
