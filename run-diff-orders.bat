@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if exist "%SCRIPT_DIR%AutoExcelKit\diff-orders.exe" (
  set "BIN=%SCRIPT_DIR%AutoExcelKit\diff-orders.exe"
) else if exist "%SCRIPT_DIR%dist\AutoExcelKit\diff-orders.exe" (
  set "BIN=%SCRIPT_DIR%dist\AutoExcelKit\diff-orders.exe"
) else (
  echo Cannot find diff-orders.exe.
  echo Please keep this file next to the AutoExcelKit folder.
  pause
  exit /b 1
)

cd /d "%SCRIPT_DIR%"
"%BIN%"

echo.
echo Processing finished. Press any key to close.
pause >nul
