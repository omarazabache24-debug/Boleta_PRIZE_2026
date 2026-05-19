@echo off
cd /d %~dp0
set HOST=127.0.0.1
set PORT=5000
python app.py
pause
