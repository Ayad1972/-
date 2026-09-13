@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo ============================================
echo تفريغ AL082026.DBF ثم نقل بيانات العلاوات
echo بدون تغيير هيكل الجدول وبنفس العربي
echo ============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
  where py >nul 2>&1
  if errorlevel 1 (
    echo Python غير موجود. شغّل أولاً:
    echo install_python.bat
    pause
    exit /b 1
  )
  set "PY=py -3"
) else (
  set "PY=python"
)

echo تثبيت المكتبات المطلوبة...
%PY% -m pip install --quiet openpyxl
if errorlevel 1 (
  echo فشل تثبيت openpyxl
  pause
  exit /b 1
)

set "EXCEL=C:\Users\ngc\Desktop\092026\NewٌRel\علاوات حسب امر   8276في 13-8-2026.xlsx"
set "DBF=C:\Users\ngc\Desktop\092026\NewٌRel\AL082026.DBF"

echo.
echo 1^) فحص الملفات...
if exist "%DBF%" if exist "%EXCEL%" (
  %PY% "%~dp0transfer_excel_to_al_dbf.py" --inspect-only --excel "%EXCEL%" --dbf "%DBF%"
) else (
  %PY% "%~dp0transfer_excel_to_al_dbf.py" --inspect-only
)
if errorlevel 1 (
  echo.
  echo لم يتم العثور على الملفات.
  echo انسخ الملفين إلى مجلد data بجانب هذا الملف ثم أعد التشغيل.
  pause
  exit /b 1
)

echo.
echo 2^) التنفيذ: تفريغ الجدول ثم النقل مع نسخة احتياطية...
if exist "%DBF%" if exist "%EXCEL%" (
  %PY% "%~dp0transfer_excel_to_al_dbf.py" --excel "%EXCEL%" --dbf "%DBF%"
) else (
  %PY% "%~dp0transfer_excel_to_al_dbf.py"
)
if errorlevel 1 (
  pause
  exit /b 1
)

echo.
echo اكتمل النقل.
pause
