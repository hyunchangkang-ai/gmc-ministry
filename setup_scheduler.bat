@echo off
:: Check for administrative privileges
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [INFO] Running as Administrator...
    goto :run
) else (
    echo [INFO] Requesting Administrative privileges...
    goto :elevate
)

:elevate
    powershell -Command "Start-Process -FilePath '%0' -Verb RunAs"
    exit /b

:run
    cd /d "%~dp0"
    python setup_scheduler.py
    echo.
    echo ==============================================
    echo Task Scheduler registration completed.
    echo Press any key to close this window.
    echo ==============================================
    pause > nul
