@echo off
echo ============================================
echo   Jig Workflow Automation - Final Setup
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install from python.org
    pause
    exit /b 1
)

echo [1/3] Creating virtual environment...
python -m venv venv
call venv\Scripts\activate.bat

echo [2/3] Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt

echo [3/3] Creating project folders...
if not exist "%USERPROFILE%\JigProjects" mkdir "%USERPROFILE%\JigProjects"

echo.
echo ============================================
echo   DONE! Now run add_to_startup.bat
echo   (Right-click it, Run as Administrator)
echo ============================================
pause