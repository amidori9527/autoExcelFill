@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if exist "%SCRIPT_DIR%add-cards.exe" (
  set "BIN=%SCRIPT_DIR%add-cards.exe"
) else if exist "%SCRIPT_DIR%AutoExcelKit\add-cards.exe" (
  set "BIN=%SCRIPT_DIR%AutoExcelKit\add-cards.exe"
) else if exist "%SCRIPT_DIR%dist\AutoExcelKit\add-cards.exe" (
  set "BIN=%SCRIPT_DIR%dist\AutoExcelKit\add-cards.exe"
) else if exist "%SCRIPT_DIR%.venv\Scripts\python.exe" (
  set "PYTHONPATH=%SCRIPT_DIR%src"
  cd /d "%SCRIPT_DIR%"
  "%SCRIPT_DIR%.venv\Scripts\python.exe" -m autoexcel.add_cards
  goto finished
) else (
  echo Cannot find add-cards.exe.
  echo Please keep this file next to the AutoExcelKit folder.
  pause
  exit /b 1
)

cd /d "%SCRIPT_DIR%"
"%BIN%"

:finished
echo.
echo Processing finished. Press any key to close.
pause >nul
