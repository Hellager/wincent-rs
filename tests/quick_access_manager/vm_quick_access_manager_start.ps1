# Session-aware launcher. No previous run directories or processes are removed.
param([Parameter(Mandatory = $true)][string]$ConfigPath)

$ErrorActionPreference = 'Stop'
$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
if ([string]$config.run_id -notmatch '^[a-zA-Z0-9-]+$') { throw 'Invalid run ID' }
$root = "C:\WincentTrack\qam-$($config.run_id)"
if ([IO.Path]::GetFullPath($ConfigPath) -ne "$root\config.json") { throw 'Config path mismatch' }
$script = "$root\guest.ps1"
$task = "Wincent-QAM-$($config.run_id)"
$user = [string]$config.guest_user
if ($user -notmatch '\\') { $user = "$env:COMPUTERNAME\$user" }
$session = (Get-Process -Id $PID).SessionId
$explorer = Get-Process explorer -ErrorAction Stop | Where-Object SessionId -eq $session
$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -ConfigPath `"$ConfigPath`""
if ($session -ne 0 -and $explorer) {
    Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -WindowStyle Hidden
} else {
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arguments
    $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Highest
    Register-ScheduledTask -TaskName $task -Action $action -Principal $principal |
        Out-Null
    Start-ScheduledTask -TaskName $task
}
Set-Content -LiteralPath "$root\started.txt" -Value $config.run_id
