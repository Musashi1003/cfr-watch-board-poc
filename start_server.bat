@echo off
setlocal
set "PYTHON=C:\Users\iec950458\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PYTHON%" (
  echo Python runtime not found: %PYTHON%
  exit /b 1
)

set "SITES_HOST=127.0.0.1"
set "SITES_PORT=8088"
set "PYTHONDONTWRITEBYTECODE=1"
echo Starting Sites prototype server at http://%SITES_HOST%:%SITES_PORT%
"%PYTHON%" -B server.py
