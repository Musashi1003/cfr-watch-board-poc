@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_act_sync.ps1"
echo ACT update finished. See .act_sync\status.json and .act_sync\act_sync.log.
pause
