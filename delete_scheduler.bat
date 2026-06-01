@echo off
title Delete Bible Auto Send Scheduler
:: Check for administrative privileges
net session >nul 2>&1
if %errorLevel% == 0 (
    goto :run
) else (
    goto :elevate
)

:elevate
    powershell -Command "Start-Process -FilePath '%0' -Verb RunAs"
    exit /b

:run
    echo [INFO] Deleting daily morning 6:00 AM scheduler task...
    schtasks /delete /tn "Bible_Kakao_Auto_Send" /f
    echo.
    echo =======================================================
    echo [SUCCESS] Daily automatic send scheduler deleted!
    echo Press any key to close this window.
    echo =======================================================
    pause > nul
