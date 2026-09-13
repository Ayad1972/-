@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo النقل يتم من داخل Visual FoxPro
echo بايثون لا يستطيع تفريغ الجدول وهو مفتوح
echo ============================================
echo.

set "DEST1=C:\Users\ngc\Desktop\092026"
set "DEST2=C:\Users\ngc\Desktop\092026\NewRel"
set "DEST3=%USERPROFILE%\Desktop\092026"

copy /Y "%~dp0transfer_alawat.prg" "%DEST1%\transfer_alawat.prg" >nul 2>&1
copy /Y "%~dp0transfer_alawat.prg" "%DEST3%\transfer_alawat.prg" >nul 2>&1
copy /Y "%~dp0transfer_alawat.prg" "%USERPROFILE%\Downloads\transfer_alawat.prg" >nul 2>&1
copy /Y "%~dp0transfer_alawat.prg" "C:\Users\ngc\Downloads\transfer_alawat.prg" >nul 2>&1

echo تم النسخ إلى:
echo %USERPROFILE%\Downloads\transfer_alawat.prg
echo.
echo لا تكتب: DO transfer_alawat.prg
echo اكتب هذا السطر كاملا:
echo.
echo DO C:\Users\ngc\Downloads\transfer_alawat.prg
echo.
echo بعد التنفيذ يجب أن يظهر:
echo 126 records
echo.
pause
