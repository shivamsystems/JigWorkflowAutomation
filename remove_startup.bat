@echo off
schtasks /delete /tn "JigWorkflowAuto" /f >nul 2>&1
echo Auto-start removed.
pause