@echo off
setlocal

cd /d "%~dp0webapp"

if not exist "package.json" (
  echo Cannot find webapp\package.json.
  pause
  exit /b 1
)

where npm.cmd >nul 2>nul
if errorlevel 1 (
  echo npm.cmd was not found. Please install Node.js first.
  pause
  exit /b 1
)

if exist "node_modules\electron\dist\electron.exe" (
  start "" /d "%cd%" "node_modules\electron\dist\electron.exe" .
  exit /b 0
) else (
  echo Installing required packages...
  call npm.cmd install
  if errorlevel 1 (
    echo.
    echo Package install failed. Please try again later.
    pause
    exit /b 1
  )

  if exist "node_modules\electron\dist\electron.exe" (
    start "" /d "%cd%" "node_modules\electron\dist\electron.exe" .
    exit /b 0
  )
)

if errorlevel 1 (
  echo.
  echo Launch failed. Please make sure Node.js is installed, then try again.
  pause
)

endlocal
