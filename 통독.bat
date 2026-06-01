@echo off
title Bible Auto Send Trigger
cd /d "%~dp0"
echo [INFO] Sending today's Bible reading message to KakaoTalk...
python send_bible_message.py --now
echo.
echo ==============================================
echo [SUCCESS] Today's message sent successfully!
echo This window will close automatically in 3 seconds.
echo ==============================================
timeout /t 3 > nul
