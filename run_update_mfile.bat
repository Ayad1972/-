@echo off
chcp 1256 >nul
setlocal
cd /d "%~dp0"

echo ============================================
echo نقل الموظفين والخدمة من Excel الى MFILE.DBF
echo مع الحفاظ على ترميز الفوكس القديم
echo ============================================

python -m pip install --quiet openpyxl dbfread dbf
if errorlevel 1 (
  echo فشل تثبيت المكتبات. تأكد من تثبيت Python.
  pause
  exit /b 1
)

echo.
echo 1^) فحص الملفات...
python "%~dp0update_mfile_from_excel.py" --inspect-only --excel "H:\New Microsoft Excel Worksheet.xlsx" --dbf "H:\MFILE.DBF"
if errorlevel 1 (
  echo.
  echo لم يتم العثور على الملفات على H:\
  echo انسخ الملفين الى مجلد data بجانب هذا الملف ثم اعد التشغيل.
  pause
  exit /b 1
)

echo.
echo 2^) تجربة بدون تعديل...
python "%~dp0update_mfile_from_excel.py" --dry-run --excel "H:\New Microsoft Excel Worksheet.xlsx" --dbf "H:\MFILE.DBF"
if errorlevel 1 (
  pause
  exit /b 1
)

echo.
set /p CONFIRM=هل تريد التنفيذ الفعلي الان؟ (Y/N): 
if /I not "%CONFIRM%"=="Y" (
  echo تم الإلغاء.
  pause
  exit /b 0
)

echo.
echo 3^) التنفيذ الفعلي مع نسخة احتياطية...
python "%~dp0update_mfile_from_excel.py" --excel "H:\New Microsoft Excel Worksheet.xlsx" --dbf "H:\MFILE.DBF"
echo.
pause
