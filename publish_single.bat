@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM  Single-file publish for GitHub Release
REM  Output: publish\Release\DocToPdfTool-vVERSION-win-x64.exe
REM          publish\Release\SHA256SUMS.txt
REM ============================================================

set VERSION=
for /f "tokens=3 delims=<>" %%v in ('findstr /c:"<Version>" DocToPdfTool.csproj') do set VERSION=%%v
if "%VERSION%"=="" set VERSION=dev

set NAME=DocToPdfTool-v%VERSION%-win-x64
set SURC=publish\SingleFile
set OUT=publish\Release

echo ==========================================
echo  Building single-file   v%VERSION%
echo ==========================================

if exist "%SURC%" rmdir /s /q "%SURC%"

dotnet publish DocToPdfTool.csproj -c Release -o "%SURC%" -p:SingleFile=true
if errorlevel 1 goto fail

if exist "%SURC%\DocToPdfTool.exe.config" del /q "%SURC%\DocToPdfTool.exe.config"
if exist "%SURC%\*.dll" del /q "%SURC%\*.dll"
if exist "%SURC%\*.xml" del /q "%SURC%\*.xml"
if exist "%SURC%\*.pdb" del /q "%SURC%\*.pdb"
if exist "%SURC%\*.lib" del /q "%SURC%\*.lib"
if exist "%SURC%\*.exp" del /q "%SURC%\*.exp"

if not exist "%OUT%" mkdir "%OUT%"
if exist "%OUT%\%NAME%.exe" del /q "%OUT%\%NAME%.exe"
copy /y "%SURC%\DocToPdfTool.exe" "%OUT%\%NAME%.exe" >nul

powershell -NoProfile -ExecutionPolicy Bypass -Command "$d='%OUT%'; $f=Get-ChildItem -Path $d -File | Where-Object { $_.Extension -eq '.exe' -or $_.Extension -eq '.zip' }; if($f){ $f | Sort-Object Name | ForEach-Object { (Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLower() + '  ' + $_.Name } | Set-Content (Join-Path $d 'SHA256SUMS.txt') -Encoding ASCII }"

call :verify
if errorlevel 1 goto fail

echo.
echo Done!
echo   %OUT%\%NAME%.exe
for %%f in ("%OUT%\%NAME%.exe") do echo   size: %%~zf bytes
echo   %OUT%\SHA256SUMS.txt
echo.
start "" "%OUT%"
goto end

REM ---- 生成后复核：重新计算并与 SHA256SUMS.txt 比对 ----
:verify
powershell -NoProfile -ExecutionPolicy Bypass -Command "$d='%OUT%'; $sums=Join-Path $d 'SHA256SUMS.txt'; $bad=0; Get-Content $sums | ForEach-Object { $p=$_ -split '\s+',2; $f=Join-Path $d $p[1].Trim(); if(-not (Test-Path $f)){ Write-Host ('  缺失: '+$p[1]); $bad++ } else { $h=(Get-FileHash $f -Algorithm SHA256).Hash.ToLower(); if($h -ne $p[0].Trim()){ Write-Host ('  校验不一致: '+$p[1]); $bad++ } } }; if($bad -gt 0){ Write-Host 'WARNING: SHA256SUMS.txt 与实际文件不一致'; exit 1 }; Write-Host '  SHA256SUMS 复核通过'; exit 0"
exit /b %errorlevel%

:fail
echo.
echo Build Failed!

:end
pause
