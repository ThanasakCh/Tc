@echo off
echo ==============================================
echo      Install Dependencies
echo ==============================================
echo Installing required libraries...
cd /d "%~dp0\.."
pip install -r requirements.txt
echo.
echo Installation Completed!
pause
