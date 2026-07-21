@echo off
REM start.bat - Training Tracker Server Launcher
REM Automatically requests admin privileges and starts server

:: Check if already running as administrator
openfiles >nul 2>&1
if errorlevel 1 (
    echo Requesting administrator privileges...
    echo.
    REM Re-run script as administrator
    powershell -Command "Start-Process cmd -ArgumentList '/c cd %cd% && python waitress_server.py && pause' -Verb RunAs"
    exit /b
)

REM If we reach here, we're running as admin
cd /d "%~dp0"
echo.
echo ========================================
echo Training Tracker Server
echo ========================================
echo Running with administrator privileges
echo.

python waitress_server.py
pause