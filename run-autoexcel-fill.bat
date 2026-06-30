@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if exist "%SCRIPT_DIR%autoexcel-fill.exe" (
  set "BIN=%SCRIPT_DIR%autoexcel-fill.exe"
) else if exist "%SCRIPT_DIR%AutoExcelKit\autoexcel-fill.exe" (
  set "BIN=%SCRIPT_DIR%AutoExcelKit\autoexcel-fill.exe"
) else if exist "%SCRIPT_DIR%dist\AutoExcelKit\autoexcel-fill.exe" (
  set "BIN=%SCRIPT_DIR%dist\AutoExcelKit\autoexcel-fill.exe"
) else (
  echo Cannot find autoexcel-fill.exe.
  echo Please keep this file next to the AutoExcelKit folder.
  pause
  exit /b 1
)

cd /d "%SCRIPT_DIR%"
"%BIN%"

echo.
echo Processing finished. Press any key to close.
pause >nul
