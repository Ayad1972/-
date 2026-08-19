@echo off
cd /d "%~dp0"
title Download MFILE_updated.DBF

set "URL=https://raw.githubusercontent.com/Ayad1972/-/main/MFILE_updated.DBF"
set "DESKTOP=%USERPROFILE%\Desktop\MFILE_updated.DBF"
set "LOCAL=%~dp0MFILE_updated.DBF"
set "WORK=%~dp0MFILE.DBF"

echo Downloading MFILE_updated.DBF ...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { Invoke-WebRequest -Uri '%URL%' -OutFile '%LOCAL%' -UseBasicParsing; Copy-Item -Force '%LOCAL%' '%DESKTOP%' -ErrorAction SilentlyContinue; if (-not (Test-Path '%WORK%')) { Copy-Item -Force '%LOCAL%' '%WORK%' }; Write-Host 'Saved:' '%LOCAL%'; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
  if exist "%LOCAL%" (
    echo Using local MFILE_updated.DBF already in this folder.
    if not exist "%WORK%" copy /Y "%LOCAL%" "%WORK%" >nul
  ) else (
    echo Download failed. Copy MFILE.DBF from the working PC into this folder.
    pause
    exit /b 0
  )
)

if exist H:\ (
  copy /Y "%LOCAL%" "H:\MFILE_updated.DBF" >nul
  echo Also copied to H:\ if the flash drive is present.
)

echo.
echo Open the file from this folder or Desktop.
explorer "%~dp0"
pause
