@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM  Directory publish for GitHub Release (zipped)
REM  Output: publish\Release\DocToPdfTool-vVERSION-win-x64.zip
REM          zip root folder contains exe + all DLLs + README.md
REM          publish\Release\SHA256SUMS.txt
REM ============================================================

set VERSION=
for /f "tokens=3 delims=<>" %%v in ('findstr /c:"<Version>" DocToPdfTool.csproj') do set VERSION=%%v
if "%VERSION%"=="" set VERSION=dev

set NAME=DocToPdfTool-v%VERSION%-win-x64
set SURC=publish\DirPublish
set STAGE=publish\Stage\%NAME%
set OUT=publish\Release

echo ==========================================
echo  Building directory publish   v%VERSION%
echo ==========================================

if exist "%SURC%" rmdir /s /q "%SURC%"
if exist "publish\Stage" rmdir /s /q "publish\Stage"

dotnet publish DocToPdfTool.csproj -c Release -o "%SURC%"
if errorlevel 1 goto fail

if exist "%SURC%\DocToPdfTool.exe.config" del /q "%SURC%\DocToPdfTool.exe.config"
if exist "%SURC%\*.xml" del /q "%SURC%\*.xml"
if exist "%SURC%\*.pdb" del /q "%SURC%\*.pdb"
if exist "%SURC%\*.lib" del /q "%SURC%\*.lib"
if exist "%SURC%\*.exp" del /q "%SURC%\*.exp"

mkdir "%STAGE%"
robocopy "%SURC%" "%STAGE%" /E /NFL /NDL /NJH /NJS /NP >nul
if exist "README.md" copy /y "README.md" "%STAGE%\README.md" >nul

if not exist "%OUT%" mkdir "%OUT%"
if exist "%OUT%\%NAME%.zip" del /q "%OUT%\%NAME%.zip"

powershell -NoProfile -ExecutionPolicy Bypass -File "build\make-zip.ps1" -SourceDir "%STAGE%" -DestinationPath "%OUT%\%NAME%.zip"
if errorlevel 1 goto fail

rmdir /s /q "publish\Stage"

powershell -NoProfile -ExecutionPolicy Bypass -File "build\make-sums.ps1" -OutputDir "%OUT%"
if errorlevel 1 goto fail

echo.
echo Done!
echo   %OUT%\%NAME%.zip
for %%f in ("%OUT%\%NAME%.zip") do echo   size: %%~zf bytes
echo   zip content: %NAME%\ (exe + all DLLs + README.md)
echo   %OUT%\SHA256SUMS.txt
echo.
start "" "%OUT%"
goto end

:fail
echo.
echo Build Failed!

:end
pause