@echo off
REM ASCII-only. Finds Excel/DBF on this PC (folder, data, Desktop, USB, H:).
cd /d "%~dp0"
title Update MFILE from Excel

call "%~dp0_find_python.bat"
if not defined PY (
  echo Python not found. Run START.bat first.
  pause
  exit /b 0
)

echo ============================================
echo Update MFILE.DBF from Excel
echo No H: drive required
echo ============================================
echo.

%PY% -m pip install --quiet openpyxl dbfread dbf
echo.
echo 1) Inspect files...
%PY% "%~dp0update_mfile_from_excel.py" --inspect-only
if errorlevel 1 (
  echo.
  echo Put MFILE.DBF and the Excel file in this folder or in data\
  echo Then run this file again.
  pause
  exit /b 0
)

echo.
echo 2) Dry run...
%PY% "%~dp0update_mfile_from_excel.py" --dry-run
if errorlevel 1 (
  pause
  exit /b 0
)

echo.
set /p CONFIRM=Apply changes now? (Y/N): 
if /I not "%CONFIRM%"=="Y" (
  echo Cancelled.
  pause
  exit /b 0
)

echo.
echo 3) Writing with backup...
%PY% "%~dp0update_mfile_from_excel.py"
echo.
pause
