@echo off
setlocal
set "PYTHON=C:\Users\iec950458\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PYTHON%" (
  echo Python runtime not found: %PYTHON%
  exit /b 1
)

set "SITES_HOST=0.0.0.0"
set "SITES_PORT=8088"
set "SITES_ALLOWED_EMPLOYEE_IDS=*"
set "PYTHONDONTWRITEBYTECODE=1"

for /f "usebackq delims=" %%I in (`powershell -NoProfile -Command "$ip=(Get-NetIPConfiguration | Where-Object { $_.IPv4Address -and $_.NetAdapter.Status -eq 'Up' } | ForEach-Object { $_.IPv4Address.IPAddress } | Select-Object -First 1); if ($ip) { $ip }"`) do set "LAN_IP=%%I"

echo Starting Sites prototype server for LAN sharing.
echo Local URL: http://127.0.0.1:%SITES_PORT%/cfr-watch
if defined LAN_IP (
  echo Coworker URL: http://%LAN_IP%:%SITES_PORT%/cfr-watch
) else (
  echo Coworker URL: http://YOUR-COMPUTER-IP:%SITES_PORT%/cfr-watch
)
echo.
echo Employee ID rule: any non-empty employee ID is accepted in this LAN prototype.
echo Keep this window open while coworkers are using the site.
echo Press Ctrl+C to stop the server.
echo.
"%PYTHON%" -B server.py
