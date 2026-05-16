@echo off
cd /d "%~dp0"
call venv\Scripts\activate.bat
python jig_workflow.py
if errorlevel 1 (
    echo.
    echo ERROR: The application crashed.
    echo Check jig_workflow.log for details.
    pause
)