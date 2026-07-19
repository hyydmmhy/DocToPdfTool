@echo off
echo Building directory publish...

dotnet publish DocToPdfTool.csproj -c Release -o publish\DirPublish

if %errorlevel% equ 0 (
    if exist "publish\DirPublish\DocToPdfTool.exe.config" del "publish\DirPublish\DocToPdfTool.exe.config"
    echo.
    echo Done! Output: publish\DirPublish\
    echo   DocToPdfTool.exe + all dependency DLLs
    start "" "publish\DirPublish"
) else (
    echo.
    echo Build Failed!
)
pause