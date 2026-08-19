@echo off
cd /d "%~dp0"
title Copy MAIN720 to Desktop

set "SRC=%~dp0MAIN720_updated.xls"
if not exist "%SRC%" set "SRC=%~dp0data\MAIN720_updated.xls"
set "DEST=%USERPROFILE%\Desktop\MAIN720_updated.xls"

if not exist "%SRC%" (
  echo File not found:
  echo %SRC%
  echo Put MAIN720_updated.xls in this folder or in data\
  pause
  exit /b 0
)

copy /Y "%SRC%" "%DEST%" >nul
if errorlevel 1 (
  echo Copy to Desktop failed.
  pause
  exit /b 0
)

echo Saved to Desktop:
echo %DEST%
explorer "%USERPROFILE%\Desktop"
pause
