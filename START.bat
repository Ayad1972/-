@echo off
REM Open Afrad on any Windows PC. ASCII only. Always exits 0.
REM Does NOT require Python. Does NOT scan A:..Z: (empty CD/card readers pop errors).
cd /d "%~dp0"
title Afrad
color 0A

if /I "%~1"=="--menu" goto MENU
if /I "%~1"=="menu" goto MENU
if /I "%~1"=="--tools" goto MENU
if /I "%~1"=="tools" goto MENU
if /I "%~1"=="--check" goto CHECK

call "%~dp0_prepare.bat" "%~dp0"

setlocal EnableDelayedExpansion
set "EXE="
set "EXEDIR=%~dp0"
for %%I in ("%EXEDIR%") do set "EXEDIR=%%~fI"
if "!EXEDIR:~-1!"=="\" if not "!EXEDIR:~-2!"==":\" set "EXEDIR=!EXEDIR:~0,-1!"

REM Known program names, then any other EXE in the folder except helpers.
for %%N in (Afrad.exe AFRAD.EXE Afrad2.exe afrad2.exe AFRAD2.EXE Nizam.exe NIZAM.EXE nizam.exe MAIN.EXE Main.exe main.exe Personnel.exe personnel.exe MFILE.exe mfile.exe FOXPRO.EXE VFP9.EXE VFP6.EXE) do (
  if not defined EXE if exist "%EXEDIR%\%%N" set "EXE=%EXEDIR%\%%N"
  if not defined EXE if exist "C:\Afrad2_work\%%N" (
    set "EXE=C:\Afrad2_work\%%N"
    set "EXEDIR=C:\Afrad2_work"
  )
  if not defined EXE if exist "%USERPROFILE%\Desktop\Afrad2_work\%%N" (
    set "EXE=%USERPROFILE%\Desktop\Afrad2_work\%%N"
    set "EXEDIR=%USERPROFILE%\Desktop\Afrad2_work"
  )
)

if not defined EXE if exist "%EXEDIR%\*.exe" (
  for %%F in ("%EXEDIR%\*.exe") do (
    set "N=%%~nxF"
    echo !N! | findstr /I "python pythonw py.exe pip setup install uninstall git winget node code" >nul
    if errorlevel 1 (
      set "EXE=%%~fF"
      set "EXEDIR=%%~dpF"
    )
  )
)

if not defined EXE if exist "C:\Afrad2_work\*.exe" (
  for %%F in ("C:\Afrad2_work\*.exe") do (
    set "N=%%~nxF"
    echo !N! | findstr /I "python pythonw py.exe pip setup install uninstall git winget node code" >nul
    if errorlevel 1 (
      set "EXE=%%~fF"
      set "EXEDIR=C:\Afrad2_work"
    )
  )
)

if defined EXE (
  for %%I in ("!EXEDIR!") do set "EXEDIR=%%~fI"
  if "!EXEDIR:~-1!"=="\" if not "!EXEDIR:~-2!"==":\" set "EXEDIR=!EXEDIR:~0,-1!"
  call "%~dp0_prepare.bat" "!EXEDIR!"
  start "" /D "!EXEDIR!" "!EXE!"
  endlocal
  exit /b 0
)

endlocal

echo.
echo ============================================
echo  Afrad folder is ready.
echo  Program EXE was not in this folder.
echo  Copy the full folder from the working PC
echo  (EXE + VFP*.DLL + DBF + CDX) then run START.bat
echo ============================================
echo.
echo Opening tools menu if Python is available...
echo.

:MENU
call "%~dp0_find_python.bat"
if defined PY (
  %PY% "%~dp0start_afrad.py" --menu
  exit /b 0
)
echo Python is optional. It is only needed for Excel tools.
echo To install it later, run install_python.bat
echo.
pause
exit /b 0

:CHECK
call "%~dp0_prepare.bat" "%~dp0"
if exist "%~dp0MFILE.DBF" echo MFILE=OK
if exist "%~dp0config.fpw" echo CONFIG=OK
if exist "%~dp0TEMP" echo TEMP=OK
echo CHECK_OK
exit /b 0
