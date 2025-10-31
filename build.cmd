@echo off
setlocal
where msbuild >nul 2>&1
if errorlevel 1 (
  echo MSBuild not found in PATH. Open "Developer Command Prompt for VS" and run build.cmd
  exit /b 1
)
msbuild CONT2BLK.csproj /t:Restore,Build /p:Configuration=Release
echo.
echo Built: bin\Release\CONT2BLK.dll
endlocal