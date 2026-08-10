@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo تثبيت Python على Windows تلقائياً
echo ============================================
echo.

net session >nul 2>&1
if errorlevel 1 (
  echo يفضّل تشغيل هذا الملف كمسؤول ^(Run as administrator^)
  echo سيتم المتابعة بصلاحيات المستخدم الحالي...
  echo.
)

where python >nul 2>&1
if not errorlevel 1 (
  echo تم العثور على Python مسبقاً:
  python --version
  echo.
  goto :INSTALL_LIBS
)

where py >nul 2>&1
if not errorlevel 1 (
  echo تم العثور على Python Launcher:
  py -3 --version
  echo.
  goto :INSTALL_LIBS
)

echo [1/3] محاولة التثبيت عبر winget...
where winget >nul 2>&1
if not errorlevel 1 (
  winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
  if not errorlevel 1 (
    echo تم التثبيت عبر winget.
    goto :REFRESH_PATH
  )
  echo فشل winget، سيتم التحميل المباشر...
) else (
  echo winget غير متوفر، سيتم التحميل المباشر...
)

echo.
echo [2/3] تحميل مثبت Python الرسمي...
set "INSTALLER=%TEMP%\python-3.12.10-amd64.exe"
set "URL=https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { Invoke-WebRequest -Uri '%URL%' -OutFile '%INSTALLER%' -UseBasicParsing; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
  echo فشل تحميل Python.
  echo حمّله يدوياً من: https://www.python.org/downloads/
  echo مهم جداً: فعّل خيار Add python.exe to PATH أثناء التثبيت.
  pause
  exit /b 1
)

echo.
echo [3/3] تثبيت Python بصمت مع إضافته إلى PATH...
"%INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_pip=1 SimpleInstall=1
if errorlevel 1 (
  echo فشل التثبيت الصامت. سيتم فتح المثبت اليدوي...
  start "" "%INSTALLER%"
  echo بعد انتهاء التثبيت تأكد من تفعيل: Add python.exe to PATH
  pause
  exit /b 1
)

:REFRESH_PATH
echo.
echo تحديث PATH للجلسة الحالية...
for /f "tokens=2*" %%A in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "USERPATH=%%B"
for /f "tokens=2*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "SYSPATH=%%B"
set "PATH=%SYSPATH%;%USERPATH%;%PATH%"

where python >nul 2>&1
if errorlevel 1 (
  where py >nul 2>&1
  if errorlevel 1 (
    echo تم التثبيت لكن الأمر python غير ظاهر بعد.
    echo أغلق هذه النافذة وافتح CMD جديد ثم نفّذ:
    echo   python --version
    pause
    exit /b 0
  )
)

:INSTALL_LIBS
echo.
echo تثبيت مكتبات مشروع نقل Excel إلى DBF...
where python >nul 2>&1
if not errorlevel 1 (
  python -m pip install --upgrade pip
  python -m pip install openpyxl dbfread dbf
  python --version
) else (
  py -3 -m pip install --upgrade pip
  py -3 -m pip install openpyxl dbfread dbf
  py -3 --version
)

echo.
echo ============================================
echo اكتمل التثبيت بنجاح
echo الخطوة التالية: شغّل run_update_mfile.bat
echo ============================================
pause
