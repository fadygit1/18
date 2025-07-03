@echo off
title Install Dependencies
color 0B

echo ========================================
echo       Install Dependencies
echo ========================================
echo.

cd /d "%~dp0"

REM Auto-detect project structure
if exist "package.json" (
    echo Found package.json in current directory
) else if exist "project\package.json" (
    echo Found package.json in project subdirectory
    cd project
) else (
    echo ERROR: package.json not found!
    echo.
    echo Make sure you have:
    echo - package.json in current directory, OR
    echo - project\package.json in subdirectory
    echo.
    pause
    exit /b 1
)

REM Check Node.js
echo Checking Node.js...
node --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Node.js is not installed!
    echo Please install Node.js from https://nodejs.org
    pause
    exit /b 1
)

echo Node.js version:
node --version
echo.

echo Cleaning old files...
if exist "node_modules" rmdir /s /q node_modules
if exist "package-lock.json" del package-lock.json

echo Installing dependencies...
echo This may take 3-7 minutes...
echo.

npm install

if errorlevel 1 (
    echo.
    echo ERROR: Installation failed!
    echo.
    echo Try these solutions:
    echo 1. Check internet connection
    echo 2. Run as administrator
    echo 3. Disable antivirus temporarily
    echo 4. Try different network
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo     INSTALLATION COMPLETED!
echo ========================================
echo.
echo You can now run the application using run.bat
echo.
pause