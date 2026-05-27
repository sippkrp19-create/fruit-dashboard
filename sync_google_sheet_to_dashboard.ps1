$ErrorActionPreference = "Stop"

$Python = "C:\Users\sippk\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$Script = Join-Path $PSScriptRoot "sync_google_sheet_to_dashboard.py"

& $Python $Script
