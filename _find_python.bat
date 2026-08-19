@echo off
REM Sets PY for the caller. Do not use setlocal here.
set "PY="

where python >nul 2>&1
if not errorlevel 1 (
  set "PY=python"
  goto :EOF
)

where py >nul 2>&1
if not errorlevel 1 (
  set "PY=py -3"
  goto :EOF
)

if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
  set "PY=%LocalAppData%\Programs\Python\Python312\python.exe"
  goto :EOF
)
if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
  set "PY=%LocalAppData%\Programs\Python\Python311\python.exe"
  goto :EOF
)
if exist "%LocalAppData%\Programs\Python\Python310\python.exe" (
  set "PY=%LocalAppData%\Programs\Python\Python310\python.exe"
  goto :EOF
)
if exist "C:\Python312\python.exe" (
  set "PY=C:\Python312\python.exe"
  goto :EOF
)
if exist "C:\Python311\python.exe" (
  set "PY=C:\Python311\python.exe"
  goto :EOF
)
if exist "%ProgramFiles%\Python312\python.exe" (
  set "PY=%ProgramFiles%\Python312\python.exe"
  goto :EOF
)

goto :EOF
