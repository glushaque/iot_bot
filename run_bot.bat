@echo off
cd /d C:\Users\1\iot_bot
set PYTHONUTF8=1

:loop
".venv\Scripts\python.exe" -u bot.py >> bot_log.txt 2>&1
timeout /t 5 /nobreak >nul
goto loop