@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo نقل المحادثات إلى OneDrive
echo ============================================
echo.
echo سيتم نسخ نسخة من المحادثات إلى مجلد ون درايف.
echo الأصل يبقى في مكانه حتى لا تتوقف البرامج.
echo.

where python >nul 2>&1
if not errorlevel 1 (
  python "%~dp0copy_chats_to_onedrive.py" %*
  set "ERR=%ERRORLEVEL%"
  goto :DONE
)

where py >nul 2>&1
if not errorlevel 1 (
  py -3 "%~dp0copy_chats_to_onedrive.py" %*
  set "ERR=%ERRORLEVEL%"
  goto :DONE
)

echo Python غير موجود. سيتم النسخ عبر PowerShell بدون تصدير الملفات المقروءة.
echo لتثبيت Python انقر مرتين على install_python.bat ثم أعد التشغيل.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0copy_chats_to_onedrive.ps1" %*
set "ERR=%ERRORLEVEL%"

:DONE
echo.
if not "%ERR%"=="0" (
  echo انتهى مع رمز: %ERR%
)
pause
exit /b %ERR%
