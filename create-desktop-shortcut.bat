@echo off
title Create Desktop Shortcut
color 0A

echo.
echo ========================================
echo    إنشاء اختصار على سطح المكتب
echo ========================================
echo.

REM الحصول على مسار سطح المكتب
for /f "tokens=3*" %%i in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v Desktop') do set DESKTOP=%%j

REM إنشاء ملف VBS لإنشاء الاختصار
echo Set oWS = WScript.CreateObject("WScript.Shell") > CreateShortcut.vbs
echo sLinkFile = "%DESKTOP%\نظام إدارة العمليات الإنشائية.lnk" >> CreateShortcut.vbs
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> CreateShortcut.vbs
echo oLink.TargetPath = "%~dp0run.bat" >> CreateShortcut.vbs
echo oLink.WorkingDirectory = "%~dp0" >> CreateShortcut.vbs
echo oLink.Description = "نظام إدارة العمليات الإنشائية" >> CreateShortcut.vbs
echo oLink.IconLocation = "%SystemRoot%\System32\shell32.dll,21" >> CreateShortcut.vbs
echo oLink.Save >> CreateShortcut.vbs

REM تشغيل ملف VBS
cscript CreateShortcut.vbs >nul

REM حذف ملف VBS المؤقت
del CreateShortcut.vbs

echo ✅ تم إنشاء الاختصار على سطح المكتب بنجاح!
echo.
echo يمكنك الآن تشغيل البرنامج من سطح المكتب
echo بالضغط على أيقونة "نظام إدارة العمليات الإنشائية"
echo.
pause