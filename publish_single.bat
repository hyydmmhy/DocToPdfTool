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

if not exist "%OUT%" mkdir "%OUT%"
if exist "%OUT%\%NAME%.exe" del /q "%OUT%\%NAME%.exe"
copy /y "%SURC%\DocToPdfTool.exe" "%OUT%\%NAME%.exe" >nul

powershell -NoProfile -ExecutionPolicy Bypass -Command "$d='%OUT%'; $f=Get-ChildItem -Path $d -File | Where-Object { $_.Extension -eq '.exe' -or $_.Extension -eq '.zip' }; if($f){ $f | Sort-Object Name | ForEach-Object { (Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLower() + '  ' + $_.Name } | Set-Content (Join-Path $d 'SHA256SUMS.txt') -Encoding ASCII }"

echo.
echo Done!
echo   %OUT%\%NAME%.exe
for %%f in ("%OUT%\%NAME%.exe") do echo   size: %%~zf bytes
echo   %OUT%\SHA256SUMS.txt
echo.
start "" "%OUT%"
goto end

:fail
echo.
echo Build Failed!

:end
pause
