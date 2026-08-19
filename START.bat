@echo off
REM ASCII-only launcher. Safe on Arabic Windows. No H: drive required.
cd /d "%~dp0"
title Afrad Personnel System
color 0A

echo ============================================
echo   AFRAD - Personnel System
echo   Nizam al-Afrad - any Windows PC
echo ============================================
echo.

call "%~dp0_find_python.bat"
if defined PY goto :RUN

echo Python not found. Installing...
call "%~dp0install_python.bat"
call "%~dp0_find_python.bat"

if not defined PY (
  echo.
  echo Python is required.
  echo 1^) Run install_python.bat
  echo 2^) Or install Python from https://www.python.org/downloads/
  echo    Enable: Add python.exe to PATH
  echo 3^) Then double-click START.bat again
  echo.
  pause
  exit /b 0
)

:RUN
%PY% "%~dp0start_afrad.py" %*
set "ERR=%ERRORLEVEL%"
if not "%ERR%"=="0" (
  echo.
  echo Finished with code %ERR%
  echo If the personnel EXE is missing, copy the full folder from the working PC.
  pause
)
exit /b %ERR%
