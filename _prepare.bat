@echo off
REM Silent path setup for FoxPro Afrad on any Windows PC.
REM No Python. No PowerShell. No codepage change. Errors discarded.
REM Optional arg 1 = application folder. Default: this script's folder.

set "APPDIR=%~1"
if "%APPDIR%"=="" set "APPDIR=%~dp0"
for %%I in ("%APPDIR%") do set "APPDIR=%%~fI"
if "%APPDIR:~-1%"=="\" if not "%APPDIR:~-2%"==":\" set "APPDIR=%APPDIR:~0,-1%"
cd /d "%APPDIR%" 2>nul

if not exist "%APPDIR%\TEMP" mkdir "%APPDIR%\TEMP" >nul 2>&1

REM Local MFILE.DBF so the program does not need a flash drive.
if not exist "%APPDIR%\MFILE.DBF" if exist "%APPDIR%\MFILE_updated.DBF" copy /Y "%APPDIR%\MFILE_updated.DBF" "%APPDIR%\MFILE.DBF" >nul 2>&1
if not exist "%APPDIR%\MFILE.DBF" if exist "%APPDIR%\data\MFILE.DBF" copy /Y "%APPDIR%\data\MFILE.DBF" "%APPDIR%\MFILE.DBF" >nul 2>&1
if not exist "%APPDIR%\MFILE.DBF" if exist "%APPDIR%\data\MFILE_updated.DBF" copy /Y "%APPDIR%\data\MFILE_updated.DBF" "%APPDIR%\MFILE.DBF" >nul 2>&1
if exist "%APPDIR%\MFILE.CDX" attrib -R "%APPDIR%\MFILE.CDX" >nul 2>&1
if exist "%APPDIR%\MFILE.DBF" attrib -R "%APPDIR%\MFILE.DBF" >nul 2>&1
if exist "%APPDIR%\data\*.DBF" attrib -R "%APPDIR%\data\*.DBF" >nul 2>&1
if exist "%APPDIR%\*.DBF" attrib -R "%APPDIR%\*.DBF" >nul 2>&1
if exist "%APPDIR%\*.CDX" attrib -R "%APPDIR%\*.CDX" >nul 2>&1
if exist "%APPDIR%\*.FPT" attrib -R "%APPDIR%\*.FPT" >nul 2>&1

REM Portable FoxPro config. Relative paths only. No COMMAND= line.
> "%APPDIR%\config.fpw" echo DEFAULT=.
>> "%APPDIR%\config.fpw" echo PATH=.;.\data;H:\;C:\Afrad2_work
>> "%APPDIR%\config.fpw" echo RESOURCE=OFF
>> "%APPDIR%\config.fpw" echo HELP=OFF
>> "%APPDIR%\config.fpw" echo SCREEN=ON
>> "%APPDIR%\config.fpw" echo TMPFILES=.\TEMP
>> "%APPDIR%\config.fpw" echo SORTWORK=.\TEMP
>> "%APPDIR%\config.fpw" echo EDITWORK=.\TEMP
>> "%APPDIR%\config.fpw" echo MVCOUNT=8192
>> "%APPDIR%\config.fpw" echo CODEPAGE=1256

REM Map H: to this folder when the letter is free (old FoxPro looks for H:\MFILE.DBF).
REM Never probe removable drive letters; empty readers show a disk popup.
subst H: "%APPDIR%" >nul 2>&1

REM Old program often hardcodes C:\Afrad2_work.
if /I not "%APPDIR%"=="C:\Afrad2_work" (
  if not exist "C:\Afrad2_work" (
    mklink /J "C:\Afrad2_work" "%APPDIR%" >nul 2>&1
  )
  if not exist "C:\Afrad2_work" mkdir "C:\Afrad2_work" >nul 2>&1
  if exist "C:\Afrad2_work" if /I not "%APPDIR%"=="C:\Afrad2_work" (
    copy /Y "%APPDIR%\*.DBF" "C:\Afrad2_work\" >nul 2>&1
    copy /Y "%APPDIR%\*.CDX" "C:\Afrad2_work\" >nul 2>&1
    copy /Y "%APPDIR%\*.FPT" "C:\Afrad2_work\" >nul 2>&1
    copy /Y "%APPDIR%\*.APP" "C:\Afrad2_work\" >nul 2>&1
    copy /Y "%APPDIR%\*.DLL" "C:\Afrad2_work\" >nul 2>&1
    copy /Y "%APPDIR%\*.EXE" "C:\Afrad2_work\" >nul 2>&1
    copy /Y "%APPDIR%\config.fpw" "C:\Afrad2_work\" >nul 2>&1
    if not exist "C:\Afrad2_work\TEMP" mkdir "C:\Afrad2_work\TEMP" >nul 2>&1
  )
)

REM Visual FoxPro runtime next to the EXE (avoids "support library" popup).
for %%D in ("%APPDIR%" "%WINDIR%\SysWOW64" "%WINDIR%\System32" "%ProgramFiles(x86)%\Common Files\Microsoft Shared\VFP" "%ProgramFiles%\Common Files\Microsoft Shared\VFP" "C:\Afrad2_work") do (
  for %%N in (VFP9R.DLL VFP9RENU.DLL VFP9RARA.DLL VFP9T.DLL VFP6R.DLL VFP6RENU.DLL VFP8R.DLL VFP8RENU.DLL VFP7R.DLL VFP7RENU.DLL MSVCR71.DLL MSVCR70.DLL gdiplus.dll) do (
    if not exist "%APPDIR%\%%N" if exist "%%~D\%%N" copy /Y "%%~D\%%N" "%APPDIR%\%%N" >nul 2>&1
  )
)

set "PATH=%APPDIR%;%PATH%"
exit /b 0
