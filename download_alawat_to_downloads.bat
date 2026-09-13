@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

set "DEST=%USERPROFILE%\Downloads\transfer_alawat.prg"
if defined USERPROFILE if exist "%USERPROFILE%\Downloads\" goto :HAVE_DEST
set "DEST=C:\Users\ngc\Downloads\transfer_alawat.prg"

:HAVE_DEST
echo ============================================
echo نسخ transfer_alawat.prg إلى Downloads
echo ============================================
echo.

if exist "%~dp0transfer_alawat.prg" (
  copy /Y "%~dp0transfer_alawat.prg" "%DEST%" >nul
  if not errorlevel 1 goto :OK
)

echo الملف غير بجانب هذا السكربت. جاري التحميل من GitHub...
set "URL=https://raw.githubusercontent.com/Ayad1972/-/cursor/transfer-al082026-excel-6bac/transfer_alawat.prg"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { Invoke-WebRequest -Uri '%URL%' -OutFile '%DEST%' -UseBasicParsing; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
  echo فشل النسخ إلى Downloads.
  pause
  exit /b 1
)

:OK
echo تم الحفظ في:
echo %DEST%
echo.
echo من Visual FoxPro Command اكتب:
echo DO "%DEST%"
echo.
explorer "%USERPROFILE%\Downloads"
pause
