@echo off
echo Building single-file...

dotnet publish DocToPdfTool.csproj -c Release -o publish\SingleFile -p:SingleFile=true

if %errorlevel% equ 0 (
    if exist "publish\SingleFile\DocToPdfTool.exe.config" del "publish\SingleFile\DocToPdfTool.exe.config"
    REM Remove build artifact DLLs — Costura embeds them into the exe
    if exist "publish\SingleFile\*.dll" del "publish\SingleFile\*.dll"
    echo.
    echo Done! Output: publish\SingleFile\DocToPdfTool.exe
    start "" "publish\SingleFile"
) else (
    echo.
    echo Build Failed!
)
pause