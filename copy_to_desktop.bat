@echo off
chcp 65001 >nul
setlocal

set "SRC=%~dp0MAIN720_updated.xls"
set "DEST=%USERPROFILE%\Desktop\MAIN720_updated.xls"

if not exist "%SRC%" (
  echo الملف غير موجود بجانب هذا الملف:
  echo %SRC%
  echo.
  echo ضع MAIN720_updated.xls بجانب copy_to_desktop.bat ثم اعد التشغيل.
  pause
  exit /b 1
)

copy /Y "%SRC%" "%DEST%" >nul
if errorlevel 1 (
  echo فشل النسخ إلى سطح المكتب.
  pause
  exit /b 1
)

echo تم الحفظ على سطح المكتب:
echo %DEST%
explorer "%USERPROFILE%\Desktop"
pause
