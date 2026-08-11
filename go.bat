@echo off
REM ASCII-only launcher - safe on Arabic Windows CMD
REM Put this file anywhere (Desktop or Downloads) and double-click.
REM Right-click -> Run as administrator  (recommended)

setlocal EnableExtensions
title PC Speed Up - Step by Step
color 0A

echo ============================================
echo   PC SPEED UP - STEP BY STEP
echo   Tasreea al-hasoob - khutwa bi khutwa
echo ============================================
echo.
echo Recommended: Run as Administrator
echo.
echo [1] Check disk free space
echo [2] Clean TEMP files
echo [3] Empty Recycle Bin
echo [4] Open Startup Apps
echo [5] High Performance power plan
echo [6] Open Disk Cleanup
echo [7] Open Task Manager
echo [8] Tips for FoxPro / DBF
echo [9] Run safe steps 1-7 now
echo [0] Exit
echo.

:MENU
set "CHOICE="
set /p CHOICE=Choose step number: 

if "%CHOICE%"=="1" goto STEP1
if "%CHOICE%"=="2" goto STEP2
if "%CHOICE%"=="3" goto STEP3
if "%CHOICE%"=="4" goto STEP4
if "%CHOICE%"=="5" goto STEP5
if "%CHOICE%"=="6" goto STEP6
if "%CHOICE%"=="7" goto STEP7
if "%CHOICE%"=="8" goto STEP8
if "%CHOICE%"=="9" goto ALL
if "%CHOICE%"=="0" goto END
echo Invalid choice.
goto MENU

:STEP1
echo.
echo === STEP 1: Disk space ===
wmic logicaldisk where "DriveType=3" get DeviceID,FreeSpace,Size
echo.
echo === Top memory processes ===
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 8 Name,@{N='MB';E={[int]($_.WorkingSet64/1MB)}} | Format-Table -AutoSize"
echo.
pause
goto MENU

:STEP2
echo.
echo === STEP 2: Cleaning TEMP ===
echo Before:
dir "%TEMP%" | find "File(s)"
del /q /f /s "%TEMP%\*" >nul 2>&1
del /q /f /s "%LOCALAPPDATA%\Temp\*" >nul 2>&1
if exist "%WINDIR%\Temp" del /q /f /s "%WINDIR%\Temp\*" >nul 2>&1
echo TEMP cleanup done. Some files in use were skipped.
echo.
pause
goto MENU

:STEP3
echo.
echo === STEP 3: Empty Recycle Bin ===
powershell -NoProfile -ExecutionPolicy Bypass -Command "Clear-RecycleBin -Force -ErrorAction SilentlyContinue"
echo Recycle Bin cleared.
echo.
pause
goto MENU

:STEP4
echo.
echo === STEP 4: Startup apps ===
echo Disable only apps you do not need at boot.
echo Keep antivirus enabled.
start "" taskmgr.exe
start "" ms-settings:startupapps
echo.
pause
goto MENU

:STEP5
echo.
echo === STEP 5: High Performance power plan ===
powercfg /SETACTIVE 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >nul 2>&1
if errorlevel 1 (
  powercfg -duplicatescheme 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >nul 2>&1
  powercfg /SETACTIVE 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >nul 2>&1
)
powercfg /GETACTIVESCHEME
echo.
pause
goto MENU

:STEP6
echo.
echo === STEP 6: Disk Cleanup ===
start "" cleanmgr.exe
echo.
pause
goto MENU

:STEP7
echo.
echo === STEP 7: Task Manager ===
start "" taskmgr.exe
echo Close heavy apps you do not need.
echo.
pause
goto MENU

:STEP8
echo.
echo === STEP 8: FoxPro / DBF tips ===
echo 1. Keep MFILE.DBF on local fast disk, not slow network.
echo 2. Close Excel before DBF update scripts.
echo 3. Do not run full antivirus scan while entering data.
echo 4. Open only one personnel-system window if possible.
echo 5. Move many Desktop icons into folders - Desktop clutter slows PC.
echo 6. After big updates, rebuild indexes inside FoxPro if available.
echo.
pause
goto MENU

:ALL
echo.
echo Running safe steps 1-7 ...
echo.
echo --- Disk space ---
wmic logicaldisk where "DriveType=3" get DeviceID,FreeSpace,Size
echo.
echo --- Clean TEMP ---
del /q /f /s "%TEMP%\*" >nul 2>&1
del /q /f /s "%LOCALAPPDATA%\Temp\*" >nul 2>&1
if exist "%WINDIR%\Temp" del /q /f /s "%WINDIR%\Temp\*" >nul 2>&1
echo TEMP done.
echo.
echo --- Recycle Bin ---
powershell -NoProfile -ExecutionPolicy Bypass -Command "Clear-RecycleBin -Force -ErrorAction SilentlyContinue"
echo Recycle Bin done.
echo.
echo --- High Performance ---
powercfg /SETACTIVE 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >nul 2>&1
powercfg /GETACTIVESCHEME
echo.
echo Opening Startup settings + Disk Cleanup + Task Manager ...
start "" taskmgr.exe
start "" ms-settings:startupapps
start "" cleanmgr.exe
echo.
echo DONE. Please RESTART the PC now.
echo Also: move Desktop files into folders to speed Explorer.
echo.
pause
goto MENU

:END
echo Bye.
endlocal
exit /b 0
