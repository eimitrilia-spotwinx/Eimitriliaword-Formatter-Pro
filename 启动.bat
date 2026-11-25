@echo off
chcp 65001 >nul
title Eimitriliaword Formatter Pro
cd /d "%~dp0"

rem Stable mode: local server + default browser (no Edge app window)
set WFP_NO_WINDOW=1

if not exist "_env\.venv\Scripts\python.exe" (
    echo [ERROR] _env\.venv\Scripts\python.exe not found.
    echo This .bat must stay next to the _env folder.
    pause
    exit /b 1
)

echo Starting Eimitriliaword Formatter Pro ...
echo Browser will open automatically. Close this window to stop the server.
echo.

"_env\.venv\Scripts\python.exe" "wfp_web.py"

echo.
echo Server stopped.
pause
