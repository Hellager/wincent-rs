# wincent refresh-sync VM experiment - guest session-1 launcher
#
# Runs in the guest through `vmrun runProgramInGuest` (session 0). It first
# cleans stale benchmark state, then creates a temporary interactive scheduled
# task so that the benchmark and Explorer share session 1, starts the task, and
# writes a run-specific "started" marker for the host. The host polls the marker
# files written by vm_refresh_sync_guest.ps1.

param(
    [string]$Root = 'C:\WincentTrack\refresh-sync-current',
    [string]$ConfigPath = 'C:\WincentTrack\refresh-sync-config.json',
    [string]$Script = 'C:\WincentTrack\run-refresh-sync.ps1',
    [string]$TaskName = 'WincentTrack-RefreshSync',
    [string]$UserName = ''
)

$ErrorActionPreference = 'Stop'

$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$runId = [string]$config.run_id
if (-not $runId) { throw 'config.run_id is missing' }
$startedMarker = "C:\WincentTrack\refresh-sync-started-$runId.txt"

if (-not $UserName) { $UserName = [string]$config.guest_user }
if (-not $UserName) { throw 'config.guest_user is missing' }
if ($UserName -notmatch '\\') { $UserName = "$env:COMPUTERNAME\$UserName" }

$explorer = Get-Process explorer -ErrorAction Stop |
    Where-Object SessionId -eq 1 |
    Select-Object -First 1
if (-not $explorer) {
    throw 'Explorer session 1 is not available'
}

# Stop any benchmark left over from a previous attempt, then delete its state.
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like '*run-refresh-sync.ps1*' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Get-Process 'wincent-refresh-sync' -ErrorAction SilentlyContinue |
    Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 800

if (Test-Path -LiteralPath $Root) {
    Remove-Item -LiteralPath $Root -Recurse -Force
}
Remove-Item -LiteralPath 'C:\WincentTrack\refresh-sync-evidence.zip' -Force -ErrorAction SilentlyContinue
Get-ChildItem -LiteralPath 'C:\WincentTrack' -Filter 'refresh-sync-started-*.txt' -File -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue
if (Test-Path -LiteralPath $Root) {
    throw "failed to clean the previous benchmark root: $Root"
}

cmd.exe /c "schtasks.exe /Delete /TN `"$TaskName`" /F >nul 2>&1"
$action = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$Script`""
cmd.exe /c "schtasks.exe /Create /TN `"$TaskName`" /TR `"$action`" /SC ONCE /ST 00:00 /RU `"$UserName`" /IT /F >nul 2>&1"
if ($LASTEXITCODE -ne 0) {
    throw "Failed to create benchmark scheduled task, exit $LASTEXITCODE"
}
cmd.exe /c "schtasks.exe /Run /TN `"$TaskName`" >nul 2>&1"
if ($LASTEXITCODE -ne 0) {
    throw "Failed to run benchmark scheduled task, exit $LASTEXITCODE"
}

(Get-Date).ToString('o') |
    Set-Content -LiteralPath $startedMarker -Encoding UTF8
