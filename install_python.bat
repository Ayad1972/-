@echo off
REM ASCII-only Python installer for Windows
cd /d "%~dp0"
title Install Python

echo ============================================
echo Install Python for Afrad tools
echo ============================================
echo.

net session >nul 2>&1
if errorlevel 1 (
  echo Optional: Run as administrator
  echo Continuing with current user rights...
  echo.
)

call "%~dp0_find_python.bat"
if defined PY (
  echo Python already found:
  %PY% --version
  echo.
  goto :INSTALL_LIBS
)

echo [1/3] Trying winget...
where winget >nul 2>&1
if not errorlevel 1 (
  winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
  if not errorlevel 1 (
    echo Installed with winget.
    goto :REFRESH_PATH
  )
  echo winget failed, downloading installer...
) else (
  echo winget not found, downloading installer...
)

echo.
echo [2/3] Downloading official Python installer...
set "INSTALLER=%TEMP%\python-3.12.10-amd64.exe"
set "URL=https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { Invoke-WebRequest -Uri '%URL%' -OutFile '%INSTALLER%' -UseBasicParsing; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
  echo Download failed.
  echo Get Python from https://www.python.org/downloads/
  echo Enable: Add python.exe to PATH
  pause
  exit /b 0
)

echo.
echo [3/3] Silent install with PATH...
"%INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_pip=1 SimpleInstall=1
if errorlevel 1 (
  echo Silent install failed. Opening installer window...
  start "" "%INSTALLER%"
  echo After setup, enable: Add python.exe to PATH
  pause
  exit /b 0
)

:REFRESH_PATH
echo.
echo Refreshing PATH...
for /f "tokens=2*" %%A in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "USERPATH=%%B"
for /f "tokens=2*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "SYSPATH=%%B"
set "PATH=%SYSPATH%;%USERPATH%;%PATH%"
call "%~dp0_find_python.bat"

if not defined PY (
  echo Python was installed but this window cannot see it yet.
  echo Close this window, open a new one, then run START.bat
  pause
  exit /b 0
)

:INSTALL_LIBS
echo.
echo Installing Afrad libraries...
%PY% -m pip install --upgrade pip
%PY% -m pip install openpyxl dbfread dbf "xlrd==1.2.0" xlwt xlutils
%PY% --version

echo.
echo ============================================
echo Setup complete. Next: double-click START.bat
echo ============================================
pause
