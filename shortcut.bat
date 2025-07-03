@echo off
title Create Desktop Shortcut
color 0A

echo ========================================
echo    Create Desktop Shortcut
echo ========================================
echo.

REM Get desktop path using multiple methods
set DESKTOP=
for /f "tokens=3*" %%i in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v Desktop 2^>nul') do set DESKTOP=%%j

REM Fallback method if first fails
if "%DESKTOP%"=="" (
    set DESKTOP=%USERPROFILE%\Desktop
)

REM Final fallback
if not exist "%DESKTOP%" (
    set DESKTOP=C:\Users\%USERNAME%\Desktop
)

echo Desktop path: %DESKTOP%

REM Check if desktop path exists
if not exist "%DESKTOP%" (
    echo ERROR: Cannot access desktop folder
    echo Tried: %DESKTOP%
    echo.
    pause
    exit /b 1
)

echo Creating shortcut...

REM Create VBS file with proper error handling
(
echo Set oWS = WScript.CreateObject^("WScript.Shell"^)
echo Set oFS = WScript.CreateObject^("Scripting.FileSystemObject"^)
echo.
echo ' Check if desktop exists
echo If Not oFS.FolderExists^("%DESKTOP%"^) Then
echo     WScript.Echo "Desktop folder not found"
echo     WScript.Quit 1
echo End If
echo.
echo ' Create shortcut
echo sLinkFile = "%DESKTOP%\Construction Management System.lnk"
echo Set oLink = oWS.CreateShortcut^(sLinkFile^)
echo oLink.TargetPath = "%~dp0run.bat"
echo oLink.WorkingDirectory = "%~dp0"
echo oLink.Description = "Construction Management System - نظام إدارة العمليات الإنشائية"
echo oLink.IconLocation = "%%SystemRoot%%\System32\shell32.dll,21"
echo.
echo ' Try to save with error handling
echo On Error Resume Next
echo oLink.Save
echo If Err.Number ^<^> 0 Then
echo     WScript.Echo "Error saving shortcut: " ^& Err.Description
echo     WScript.Quit 1
echo Else
echo     WScript.Echo "Shortcut created successfully"
echo End If
) > CreateShortcut.vbs

REM Run VBS file and capture output
echo Running shortcut creation script...
for /f "delims=" %%i in ('cscript //NoLogo CreateShortcut.vbs 2^>^&1') do (
    echo %%i
    if "%%i"=="Shortcut created successfully" set SUCCESS=1
)

REM Clean up
if exist "CreateShortcut.vbs" del CreateShortcut.vbs

REM Check result
if defined SUCCESS (
    echo.
    echo ========================================
    echo     SHORTCUT CREATED SUCCESSFULLY!
    echo ========================================
    echo.
    echo You can now run the application from desktop
    echo by double-clicking "Construction Management System"
    echo.
    echo Location: %DESKTOP%\Construction Management System.lnk
) else (
    echo.
    echo ========================================
    echo        SHORTCUT CREATION FAILED
    echo ========================================
    echo.
    echo Possible solutions:
    echo 1. Run this script as Administrator
    echo 2. Check if desktop folder is accessible
    echo 3. Try creating shortcut manually
    echo.
    echo Manual steps:
    echo 1. Right-click on desktop
    echo 2. Select "New" ^> "Shortcut"
    echo 3. Browse to: %~dp0run.bat
    echo 4. Name it: Construction Management System
)

echo.
pause