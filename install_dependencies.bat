@echo off
echo ========================================
echo Outlook Email Backup Tool - Installation
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

echo Python found!
python --version
echo.

echo Installing required dependencies...
echo This may take a few minutes...
echo.

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip

echo.
echo Installing packages from requirements.txt...
python -m pip install -r requirements.txt

echo.
if errorlevel 1 (
    echo.
    echo ERROR: Installation failed.
    echo Please check your internet connection and try again.
    pause
    exit /b 1
) else (
    echo ========================================
    echo Installation completed successfully!
    echo ========================================
    echo.
    echo You can now run the application by:
    echo   1. Double-clicking run_backup_tool.bat
    echo   2. Or running: python main.py
    echo.
    pause
)
