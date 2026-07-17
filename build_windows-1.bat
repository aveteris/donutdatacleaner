@echo off
REM ============================================================
REM Build script: Donut Data Cleaner -> ONE standalone .exe
REM No terminal window will appear when the final app runs.
REM Run this on a Windows machine with Python installed.
REM ============================================================

echo Installing dependencies...
pip install --upgrade pip
pip install pandas openpyxl pyinstaller

echo Building single-file windowed executable...
pyinstaller --onefile --windowed --icon=app_icon.ico --name=DonutDataCleaner donut_data_cleaner.py

echo.
echo ============================================================
echo Done! Your app is a single file here:
echo    dist\DonutDataCleaner.exe
echo Double-click it to run - no terminal window, no Python needed.
echo ============================================================
pause
