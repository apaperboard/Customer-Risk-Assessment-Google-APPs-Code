@echo off
setlocal
set ROOT=%~dp0
set DIST_DIR=%ROOT%dist
set PORT=5173

if not exist "%DIST_DIR%\index.html" (
  echo Build not found.
  echo Please ask the owner to run: npm install && npm run build
  echo Once the dist folder exists, run this file again.
  pause
  exit /b 1
)

REM Start lightweight local server via PowerShell (no installs needed)
start "AR-Dashboard Server" powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%serve.ps1" -Port %PORT% -Root "%DIST_DIR%"
timeout /t 1 >nul
start "" http://localhost:%PORT%/
