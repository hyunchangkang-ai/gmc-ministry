@echo off
cd /d "%~dp0"
echo =======================================
echo     Antigravity GitHub Auto Sync
echo =======================================
echo.

echo 1. Pulling latest changes...
git pull origin main
echo.

echo 2. Staging new changes...
git add .
echo.

echo 3. Committing changes...
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set datetime=%%I
set sync_time=%datetime:~0,4%-%datetime:~4,2%-%datetime:~6,2% %datetime:~8,2%:%datetime:~10,2%
git commit -m "Auto sync from Windows: %sync_time%"
echo.

echo 4. Pushing to GitHub...
git push origin main
echo.

echo =======================================
echo     Sync Completed!
echo =======================================
timeout /t 5
