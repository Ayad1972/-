@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo ============================================
echo تنظيف استمارة مخصصات.pdf طبق الاصل
echo ازالة النقاط السود وخط اليد
echo ============================================
echo.

set "SRC=C:\Users\ngc\Desktop\092026\مخصصات.pdf"
if not exist "%SRC%" (
  echo الملف الأصلي غير موجود:
  echo %SRC%
  pause
  exit /b 1
)

where python >nul 2>&1
if errorlevel 1 (
  where py >nul 2>&1
  if errorlevel 1 (
    echo Python غير موجود. شغّل أولاً install_python.bat
    pause
    exit /b 1
  )
  set "PY=py -3"
) else (
  set "PY=python"
)

%PY% -m pip install --quiet opencv-python-headless numpy pymupdf
%PY% "%~dp0clean_original_makhsasat.py" "%SRC%"
if errorlevel 1 (
  pause
  exit /b 1
)

echo.
echo الناتج:
echo مخصصات_مطبوعة.pdf
explorer "%USERPROFILE%\Downloads"
pause
