@echo off
REM start.bat - Training Tracker Server Launcher

cd /d "%~dp0"
echo.
echo ========================================
echo Training Tracker Server
echo ========================================
echo.

python waitress_server.py
pause