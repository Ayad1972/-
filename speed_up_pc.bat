@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo تسريع الحاسوب خطوة بخطوة
echo ============================================
echo.
echo يفضّل التشغيل كمسؤول لبعض الخطوات ^(تنظيف Temp النظام / chkdsk^)
echo يمكنك المتابعة بصلاحيات عادية أيضاً.
echo.

REM إن مُرّر رقم خطوة: نفّذها مباشرة
if not "%~1"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0speed_up_pc.ps1" -Step "%~1"
  exit /b %ERRORLEVEL%
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0speed_up_pc.ps1"
set "ERR=%ERRORLEVEL%"
echo.
if not "%ERR%"=="0" (
  echo حدث خطأ أثناء التنفيذ. رمز الخروج: %ERR%
  pause
  exit /b %ERR%
)
exit /b 0
