@echo off
REM Exact duplicate inventory - report first, no delete unless APPLY
setlocal EnableExtensions
cd /d "%~dp0"
title Exact Duplicates - Report First

where python >nul 2>&1
if errorlevel 1 (
  where py >nul 2>&1
  if errorlevel 1 (
    echo Python not found. Run install_python.bat first.
    pause
    exit /b 1
  )
  set "PY=py -3"
) else (
  set "PY=python"
)

echo ============================================
echo  Exact duplicates 100%% - REPORT ONLY
echo  Keep newest / last copy later with APPLY
echo ============================================
echo.
echo Scans Desktop + Downloads by default.
echo.

if /I "%~1"=="APPLY" goto APPLY
if /I "%~1"=="--apply" goto APPLY

echo Step 1: creating report only...
%PY% "%~dp0find_exact_duplicates.py"
set "ERR=%ERRORLEVEL%"
echo.
if not "%ERR%"=="0" (
  echo Failed. Code=%ERR%
  pause
  exit /b %ERR%
)
echo.
echo Review the report on your Desktop:
echo   duplicates_report_*.txt
echo   duplicates_report_*.csv
echo.
echo If the report is correct, run again with APPLY:
echo   find_dupes.bat APPLY
echo.
pause
exit /b 0

:APPLY
echo Step 2: MOVE duplicates to quarantine, keep newest only...
echo Nothing is permanently deleted.
echo.
%PY% "%~dp0find_exact_duplicates.py" --apply
set "ERR=%ERRORLEVEL%"
echo.
if not "%ERR%"=="0" (
  echo Failed. Code=%ERR%
  pause
  exit /b %ERR%
)
echo Done. Check Desktop\_DUPLICATES_QUARANTINE
echo Restart PC after big cleanups.
pause
exit /b 0
