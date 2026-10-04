# Runs tools/ranman_sync.py for the Windows scheduled task «Ranman sync»
# (weekly, Mondays) and appends what it did to
# %LOCALAPPDATA%\ranman_sync\sync.log. Extra arguments go to the sync, e.g.
#   powershell -File tools\ranman_sync_weekly.ps1 --dry-run
# The task runs as the signed-in user; the URL and token are read from that
# user's environment (setx RANMAN_SYNC_URL / RANMAN_SYNC_TOKEN), which a task
# started at a fixed time does not always inherit.
$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot
foreach ($v in 'RANMAN_SYNC_URL', 'RANMAN_SYNC_TOKEN') {
    if (-not [Environment]::GetEnvironmentVariable($v, 'Process')) {
        [Environment]::SetEnvironmentVariable($v, [Environment]::GetEnvironmentVariable($v, 'User'), 'Process')
    }
}
$env:PYTHONIOENCODING = 'utf-8'
# read the sync's output as UTF-8 too, or the log gets «┬╖» for «·»
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$dir = Join-Path $env:LOCALAPPDATA 'ranman_sync'
New-Item -ItemType Directory -Force $dir | Out-Null
$log = Join-Path $dir 'sync.log'
$python = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $python) { $python = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python311\python.exe' }

"==== $(Get-Date -Format 'yyyy-MM-dd HH:mm') ranman_sync $($args -join ' ')" | Out-File -FilePath $log -Append -Encoding utf8
Push-Location $repo
try {
    & $python (Join-Path $repo 'tools\ranman_sync.py') @args 2>&1 | ForEach-Object { "$_" } | Out-File -FilePath $log -Append -Encoding utf8
    $code = $LASTEXITCODE
} finally {
    Pop-Location
}
"==== exit $code" | Out-File -FilePath $log -Append -Encoding utf8
exit $code
