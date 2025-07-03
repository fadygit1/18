@echo off
title Construction Management System
color 0A
chcp 65001 >nul

echo ========================================
echo    Construction Management System
echo ========================================
echo.

cd /d "%~dp0"

REM Auto-detect project structure
set PROJECT_DIR=
if exist "src\main.tsx" (
    set PROJECT_DIR=%~dp0
    echo Found project files in current directory
) else if exist "project\src\main.tsx" (
    set PROJECT_DIR=%~dp0project
    echo Found project files in project subdirectory
    cd project
) else (
    echo ERROR: Cannot find project files!
    echo.
    echo Looking for:
    echo - src\main.tsx in current directory
    echo - project\src\main.tsx in subdirectory
    echo.
    echo Current directory: %~dp0
    echo.
    pause
    exit /b 1
)

REM Check Node.js
echo Checking Node.js...
node --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Node.js is not installed!
    echo.
    echo Please download and install Node.js from:
    echo https://nodejs.org
    echo.
    pause
    exit /b 1
)

echo Node.js found: 
node --version

REM Check package.json
if not exist "package.json" (
    echo ERROR: package.json not found in project directory!
    echo Project directory: %PROJECT_DIR%
    echo.
    pause
    exit /b 1
)

REM Install dependencies if needed
if not exist "node_modules" (
    echo Installing dependencies...
    echo This may take 2-5 minutes...
    echo.
    npm install
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies!
        echo.
        pause
        exit /b 1
    )
    echo Dependencies installed successfully!
)

echo.
echo ========================================
echo        STARTING APPLICATION
echo ========================================
echo.
echo Browser will open automatically at:
echo http://localhost:5173
echo.
echo WARNING: Do not close this window while using the app!
echo.

REM Wait 3 seconds then open browser
timeout /t 3 /nobreak >nul
start "" "http://localhost:5173"

REM Start the application
npm run dev

echo.
echo Application closed.
pause