@echo off
cd /d "%~dp0"
title Find PNO in flash mfile.dbf

set "PNO=396176"
if not "%~1"=="" set "PNO=%~1"

set "OUT=%USERPROFILE%\Desktop\pno_%PNO%_folders.txt"
set "OUT2=H:\pno_%PNO%_folders.txt"

echo ============================================
echo Find employee PNO=%PNO% in mfile.dbf folders
echo ============================================
echo.

set "PY="
where python >nul 2>&1 && set "PY=python"
if "%PY%"=="" where py >nul 2>&1 && set "PY=py -3"

if not "%PY%"=="" (
  %PY% "%~dp0find_pno_in_flash.py" --pno "%PNO%" --root "H:\" --root "%~dp0" --root "%~dp0data" --out "%OUT%"
  if exist "H:\" copy /Y "%OUT%" "%OUT2%" >nul 2>&1
  echo.
  echo Report:
  echo %OUT%
  pause
  exit /b 0
)

echo Python not found. Using PowerShell scan...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$pno='%PNO%'; $needle=[Text.Encoding]::ASCII.GetBytes($pno); $roots=@(); if (Test-Path 'H:\') { $roots += 'H:\' }; $roots += '%~dp0'; $hits=@(); foreach ($root in $roots) { Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^(mfile|MFILE)(_updated)?\.(dbf|DBF)$' } | ForEach-Object { try { $b=[IO.File]::ReadAllBytes($_.FullName); $found=$false; for ($i=0; $i -le $b.Length-$needle.Length; $i++) { $ok=$true; for ($j=0; $j -lt $needle.Length; $j++) { if ($b[$i+$j] -ne $needle[$j]) { $ok=$false; break } } ; if ($ok) { $found=$true; break } } ; if ($found) { $hits += $_.DirectoryName; Write-Host ('FOUND: '+$_.DirectoryName+'  ['+$_.Name+']') } } catch {} } } ; Write-Host ('folders='+$hits.Count); $out='%OUT%'; $txt=@('PNO='+$pno; 'FOLDERS:'; $hits); Set-Content -Path $out -Value $txt -Encoding UTF8; if (Test-Path 'H:\') { Copy-Item -Force $out 'H:\pno_%PNO%_folders.txt' -ErrorAction SilentlyContinue }; Write-Host ('Report: '+$out)"

echo.
pause
