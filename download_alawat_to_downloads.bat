@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

set "SRC=%~dp0transfer_alawat.prg"
set "DL=%USERPROFILE%\Downloads\transfer_alawat.prg"
set "URL=https://raw.githubusercontent.com/Ayad1972/-/cursor/transfer-al082026-excel-6bac/transfer_alawat.prg"

echo ============================================
echo وضع transfer_alawat.prg في Downloads
echo وبجانب جدول الفوكس
echo ============================================
echo.

if not exist "%USERPROFILE%\Downloads\" mkdir "%USERPROFILE%\Downloads\" >nul 2>&1

if exist "%SRC%" (
  copy /Y "%SRC%" "%DL%" >nul
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try { Invoke-WebRequest -Uri '%URL%' -OutFile '%DL%' -UseBasicParsing; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
  if errorlevel 1 (
    echo فشل التحميل.
    pause
    exit /b 1
  )
)

if not exist "%DL%" (
  echo الملف غير موجود في Downloads.
  pause
  exit /b 1
)

copy /Y "%DL%" "C:\Users\ngc\Downloads\transfer_alawat.prg" >nul 2>&1
copy /Y "%DL%" "%USERPROFILE%\Desktop\092026\transfer_alawat.prg" >nul 2>&1
copy /Y "%DL%" "C:\Users\ngc\Desktop\092026\transfer_alawat.prg" >nul 2>&1
copy /Y "%DL%" "C:\Users\ngc\Desktop\092026\NewRel\transfer_alawat.prg" >nul 2>&1

for /d %%D in ("%USERPROFILE%\Desktop\092026\*") do (
  copy /Y "%DL%" "%%D\transfer_alawat.prg" >nul 2>&1
  if exist "%%D\AL082026.DBF" copy /Y "%DL%" "%%D\transfer_alawat.prg" >nul 2>&1
)

echo تم الحفظ في:
echo %DL%
echo.
echo لا تكتب: DO transfer_alawat.prg
echo اكتب هذا السطر كاملا في Command:
echo.
echo DO C:\Users\ngc\Downloads\transfer_alawat.prg
echo.
explorer "%USERPROFILE%\Downloads"
pause
