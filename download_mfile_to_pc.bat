@echo off
chcp 65001 >nul
setlocal

set "URL=https://raw.githubusercontent.com/Ayad1972/-/main/MFILE_updated.DBF"
set "DESKTOP=%USERPROFILE%\Desktop\MFILE_updated.DBF"
set "FLASH=H:\MFILE_updated.DBF"

echo جاري تحميل MFILE_updated.DBF إلى حاسبتك...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { Invoke-WebRequest -Uri '%URL%' -OutFile '%DESKTOP%' -UseBasicParsing; Write-Host 'تم الحفظ على سطح المكتب:' '%DESKTOP%'; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
  echo فشل التحميل. تأكد من الإنترنت ثم أعد المحاولة.
  pause
  exit /b 1
)

if exist H:\ (
  copy /Y "%DESKTOP%" "%FLASH%" >nul
  if not errorlevel 1 (
    echo تم نسخه أيضاً إلى الفلاش:
    echo %FLASH%
  )
)

echo.
echo افتح الملف من سطح المكتب أو من H:\
explorer "%USERPROFILE%\Desktop"
pause
