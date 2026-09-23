@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo ============================================
echo إنشاء استمارة مخصصات جديدة بنفس القياس
echo ============================================
echo.

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

%PY% -m pip install --quiet reportlab arabic-reshaper python-bidi pymupdf pypdf
%PY% "%~dp0recreate_makhsasat_form.py"
if errorlevel 1 (
  pause
  exit /b 1
)

echo.
echo افتح الملف من Desktop\092026 أو من Downloads باسم:
echo مخصصات_جديدة.pdf
explorer "%USERPROFILE%\Downloads"
pause
