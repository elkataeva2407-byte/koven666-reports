@echo off
setlocal
cd /d "%~dp0"

if not exist ".env" (
  echo ERROR: .env not found. Copy .env.example to .env and fill BOT_TOKEN.
  pause
  exit /b 1
)

for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
  if not "%%A"=="" if not "%%A:~0,1%%"=="#" set "%%A=%%B"
)

if "%BOT_TOKEN%"=="" (
  echo ERROR: BOT_TOKEN is empty in .env
  pause
  exit /b 1
)
if "%ADMIN_ID%"=="" (
  echo ERROR: ADMIN_ID is empty in .env
  pause
  exit /b 1
)

echo ===================================
echo KOVEN 666 - AUTO MODE
echo Excel -> HTML automatically
 echo Press Ctrl+C to stop.
echo ===================================

:restart
python bot.py
set ERR=%ERRORLEVEL%
echo Bot stopped with code %ERR%. Restarting in 5 seconds...
timeout /t 5 /nobreak >nul
goto restart
