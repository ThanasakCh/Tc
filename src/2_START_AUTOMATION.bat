@echo off
echo ==============================================
echo        EC ^& EAR AUTOMATION WORKFLOW
echo ==============================================
echo Starting AI Processing...
echo.

cd /d "%~dp0\.."
python src\pipeline.py

echo.
echo ==============================================
echo Process Completed! Check Data_layout folder.
echo ==============================================
pause
