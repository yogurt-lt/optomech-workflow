@echo off
cd /d "%~dp0"
py -3 --version >nul 2>&1
if not errorlevel 1 (
    py -3 server.py
    goto end
)
python --version >nul 2>&1
if not errorlevel 1 (
    python server.py
    goto end
)
echo Install Python 3.10 or newer, then run this file again.
:end
pause
