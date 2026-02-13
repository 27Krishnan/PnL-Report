@echo off
title P^&L Report Launcher
echo Starting P^&L Report Local Server...

:: Check if Node.js is installed
node -v >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not installed! Please install it from https://nodejs.org/
    pause
    exit /b
)

:: Navigate to the directory of this batch file
cd /d "%~dp0"

echo Opening browser at http://localhost:3000...
start "" "http://localhost:3000"

echo.
echo --------------------------------------------------
echo SERVER IS RUNNING. 
echo KEEP THIS WINDOW OPEN WHILE USING THE APP.
echo --------------------------------------------------
echo.

npx serve -l 3000 .
