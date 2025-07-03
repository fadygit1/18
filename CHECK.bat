@echo off
title System Check
color 0E

echo ========================================
echo           System Check
echo ========================================
echo.

echo Computer: %COMPUTERNAME%
echo User: %USERNAME%
echo Script location: %~dp0
echo.

REM Check project structure
echo Checking project structure...
if exist "src\main.tsx" (
    echo [OK] Found src\main.tsx in current directory
    set PROJECT_OK=1
) else if exist "project\src\main.tsx" (
    echo [OK] Found project\src\main.tsx in subdirectory
    set PROJECT_OK=1
) else (
    echo [X] Project files not found
    echo     Looking for src\main.tsx
    set PROJECT_OK=0
)

if exist "package.json" (
    echo [OK] Found package.json in current directory
) else if exist "project\package.json" (
    echo [OK] Found project\package.json in subdirectory
) else (
    echo [X] package.json not found
)

REM Check Node.js
echo.
echo Checking Node.js...
node --version >nul 2>&1
if errorlevel 1 (
    echo [X] Node.js: NOT INSTALLED
    echo     Download from: https://nodejs.org
) else (
    for /f "tokens=*" %%i in ('node --version') do echo [OK] Node.js: %%i
)

REM Check npm
npm --version >nul 2>&1
if errorlevel 1 (
    echo [X] npm: NOT AVAILABLE
) else (
    for /f "tokens=*" %%i in ('npm --version') do echo [OK] npm: %%i
)

REM Check dependencies
echo.
echo Checking dependencies...
if exist "node_modules" (
    echo [OK] Dependencies installed in current directory
) else if exist "project\node_modules" (
    echo [OK] Dependencies installed in project subdirectory
) else (
    echo [X] Dependencies: NOT INSTALLED
    echo     Run install.bat to fix this
)

echo.
echo ========================================
echo         Recommendations
echo ========================================
echo.

if %PROJECT_OK%==0 (
    echo 1. Make sure project files are in correct location
    echo    - src\main.tsx should exist
    echo    - package.json should exist
    echo.
)

node --version >nul 2>&1
if errorlevel 1 (
    echo 2. Install Node.js from https://nodejs.org
    echo.
)

echo 3. Run install.bat if dependencies are missing
echo 4. Run run.bat to start the application

echo.
pause