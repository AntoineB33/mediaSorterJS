@echo off
REM Attempt to kill any Node.js process running on port 3000
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :3000') do (
    echo Killing process %%a
    taskkill /f /pid %%a
)
pause
