@echo off
title Quick Start
color 0C

echo ========================================
echo          Quick Start
echo ========================================
echo.

cd /d "%~dp0"

REM Quick project detection
if exist "src\main.tsx" (
    echo Starting from current directory...
) else if exist "project\src\main.tsx" (
    echo Starting from project subdirectory...
    cd project
) else (
    echo ERROR: Project files not found!
    pause
    exit /b 1
)

echo Opening browser...
start "" "http://localhost:5173"

echo Starting development server...
npm run dev

pause