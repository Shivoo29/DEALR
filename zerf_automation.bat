@echo off
REM ZERF Automation System - Windows Startup Script
REM This script runs the ZERF automation system at Windows startup

REM Change to the script directory
cd /d "%~dp0"

REM Set up environment
set SCRIPT_DIR=%~dp0
set PYTHON_SCRIPT=%SCRIPT_DIR%zerf_automation_system.py
set LOG_FILE=%SCRIPT_DIR%logs\startup.log

REM Create logs directory if it doesn't exist
if not exist "%SCRIPT_DIR%logs" mkdir "%SCRIPT_DIR%logs"

REM Log startup attempt
echo [%date% %time%] Starting ZERF Automation System >> "%LOG_FILE%"

REM Check if Python script exists
if not exist "%PYTHON_SCRIPT%" (
    echo [%date% %time%] ERROR: Python script not found at %PYTHON_SCRIPT% >> "%LOG_FILE%"
    exit /b 1
)

REM Start the automation system in background mode
echo [%date% %time%] Launching automation system in background mode >> "%LOG_FILE%"
python "%PYTHON_SCRIPT%" --background >> "%LOG_FILE%" 2>&1

REM Log completion
echo [%date% %time%] ZERF Automation System startup complete >> "%LOG_FILE%"