@echo off
echo Starting CalCO...

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from python.org
    pause
    exit /b 1
)

REM Install required packages
cd "%~dp0server"
echo Installing required packages...
python -m pip install -r requirements.txt

REM Start the server
echo Starting server...
start python new.py

REM Wait for server to initialize
echo Waiting for server to start...
timeout 5

REM Start the Flutter application
cd "%~dp0Release"
echo Starting CalCO application...
start coflut.exe

echo CalCO is now running!
pause