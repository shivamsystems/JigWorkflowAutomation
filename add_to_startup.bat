@echo off
net session >nul 2>&1
if errorlevel 1 (
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)
cd /d "%~dp0"
schtasks /delete /tn "JigWorkflowAuto" /f >nul 2>&1
schtasks /create /tn "JigWorkflowAuto" /tr "wscript.exe \"%~dp0run_silent.vbs\"" /sc onlogon /rl highest /f
echo.
echo SUCCESS! Tool will auto-start with Windows.
pause