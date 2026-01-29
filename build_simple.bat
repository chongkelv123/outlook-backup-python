@echo off
echo ================================================
echo Simple Build - One Command
echo ================================================
echo.
echo Installing PyInstaller (if needed)...
pip install pyinstaller
echo.
echo Building executable...
echo This may take a few minutes...
echo.

pyinstaller --name "OutlookBackupTool" --onefile --windowed --clean --add-data "config.json;." --hidden-import win32com --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32timezone --hidden-import tkcalendar --hidden-import babel.numbers main.py

echo.
echo ================================================
echo Build Complete!
echo ================================================
echo.
echo Your executable is at: dist\OutlookBackupTool.exe
echo.
echo You can test it now:
echo   cd dist
echo   OutlookBackupTool.exe
echo.
pause
