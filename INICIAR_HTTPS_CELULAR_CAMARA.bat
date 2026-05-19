@echo off
cd /d %~dp0
set APP_SSL=1
set HOST=0.0.0.0
set PORT=5000
python app.py
pause
