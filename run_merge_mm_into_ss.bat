@echo off
cd /d "%~dp0"
title Merge mm.xlsx into ss.xls

call "%~dp0_find_python.bat"
if not defined PY (
  echo Python not found. Run START.bat first.
  pause
  exit /b 0
)

echo ============================================
echo Merge numbers from mm.xlsx into ss.xls
echo Files are searched in this folder, data\, Desktop, USB
echo ============================================
echo.

%PY% -m pip install --quiet openpyxl "xlrd==1.2.0" >nul 2>&1
%PY% "%~dp0merge_mm_into_ss.py"
echo.
pause
