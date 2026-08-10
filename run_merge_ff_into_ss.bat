@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo دمج ارقام ff.xlsx داخل ss.xls حسب الاسم
echo ============================================
echo.
echo المدخلات:
echo   H:\ss.xls
echo   H:\ff.xlsx
echo الناتج:
echo   H:\ss_updated.xls
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0merge_mm_into_ss.ps1" -SsPath "H:\ss.xls" -MmPath "H:\ff.xlsx" -OutPath "H:\ss_updated.xls"
echo.
pause
