$ErrorActionPreference = "Stop"
$workerRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $workerRoot

# The scheduled mailer already uses this Python installation and Outlook profile.
& py (Join-Path $workerRoot "act_sync.py") @args
exit $LASTEXITCODE
