# Host: compile only, transfer, orchestrate and collect. NEVER run wincent here.
[CmdletBinding()]
param(
    [string]$Vmx = '',
    [string]$Vmrun = '',
    [string]$VmPassword = '',
    [string]$GuestUser = '',
    [string]$GuestPassword = '',
    [ValidateSet('', 'win10', 'win11')][string]$Platform = '',
    [string]$EvidenceDir = '',
    [string]$EnvFile = '',
    [ValidateRange(1200, 7200)][int]$TimeoutSeconds = 1500
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $EnvFile) { $EnvFile = Join-Path $PSScriptRoot '.env' }
$settings = @{}
if (Test-Path -LiteralPath $EnvFile) {
    foreach ($line in Get-Content -LiteralPath $EnvFile) {
        $line = $line.Trim()
        if (-not $line -or $line.StartsWith('#')) { continue }
        $parts = $line.Split('=', 2)
        if ($parts.Count -ne 2) { throw 'Invalid .env entry (expected key=value)' }
        $settings[$parts[0].Trim()] = $parts[1].Trim().Trim('"').Trim("'")
    }
}
function Setting([string]$Name) {
    if ($settings.ContainsKey($Name)) { return [string]$settings[$Name] }
    return [Environment]::GetEnvironmentVariable($Name)
}
if (-not $Vmx) { $Vmx = Setting 'WINCENT_VM_VMX' }
if (-not $Vmrun) { $Vmrun = Setting 'WINCENT_VM_VMRUN' }
if (-not $VmPassword) { $VmPassword = Setting 'WINCENT_VM_PASSWORD' }
if (-not $GuestUser) { $GuestUser = Setting 'WINCENT_VM_GUEST_USER' }
if (-not $GuestPassword) { $GuestPassword = Setting 'WINCENT_VM_GUEST_PASSWORD' }
if (-not $Platform) { $Platform = Setting 'WINCENT_VM_PLATFORM' }
if (-not $Vmx -or -not $Vmrun -or -not $GuestUser -or -not $GuestPassword) {
    throw 'VM path, vmrun path and guest credentials are required; see .env.example'
}
if ($Platform -notin @('win10', 'win11')) { throw 'Specify -Platform win10 or win11 explicitly' }
$Vmx = (Resolve-Path -LiteralPath $Vmx).Path
$Vmrun = (Resolve-Path -LiteralPath $Vmrun).Path
$RunId = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
if (-not $EvidenceDir) { $EvidenceDir = "temp\track_wincent\evidence\qam-$Platform-$RunId" }
if (-not [IO.Path]::IsPathRooted($EvidenceDir)) { $EvidenceDir = Join-Path $RepoRoot $EvidenceDir }
$EvidenceDir = [IO.Path]::GetFullPath($EvidenceDir)
if (Test-Path -LiteralPath $EvidenceDir) { throw "Evidence directory already exists: $EvidenceDir" }
New-Item -ItemType Directory -Path $EvidenceDir | Out-Null
$GuestRoot = "C:\WincentTrack\qam-$RunId"
$Snapshot = "wincent-qam-$RunId"
$snapshotCreated = $false
$rollback = $false
$failure = $null
$wasRunning = $false

function Vm([string[]]$Arguments, [switch]$Guest, [switch]$AllowFailure) {
    $vmArgs = @('-T', 'ws')
    if ($VmPassword) { $vmArgs += @('-vp', $VmPassword) }
    if ($Guest) { $vmArgs += @('-gu', $GuestUser, '-gp', $GuestPassword) }
    $output = & $Vmrun @vmArgs @Arguments 2>&1
    $code = $LASTEXITCODE
    $text = ($output | Out-String).Trim()
    # Never include credentials or full argument lists in errors/transcripts.
    if ($code -ne 0 -and -not $AllowFailure) {
        throw "VMware operation '$($Arguments[0])' failed ($code): $text"
    }
    [pscustomobject]@{ code = $code; text = $text }
}

function Guest-Exists([string]$Path) {
    (Vm @('fileExistsInGuest', $Vmx, $Path) -Guest -AllowFailure).code -eq 0
}

function Pull-Evidence {
    foreach ($name in @(
        'environment.json', 'checks.tsv', 'result.json', 'final-state.json', 'test.log',
        'FAILURE.txt', 'RESTORE_FAILURE.txt', 'ARTIFACT_FAILURE.txt', 'evidence.zip'
    )) {
        if (Guest-Exists "$GuestRoot\$name") {
            Vm @('copyFileFromGuestToHost', $Vmx, "$GuestRoot\$name",
                (Join-Path $EvidenceDir $name)) -Guest | Out-Null
        }
    }
}

function Verify-Archive {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead("$EvidenceDir\evidence.zip")
    $hasher = [Security.Cryptography.SHA256]::Create()
    try {
        $entries = @{}
        foreach ($entry in $zip.Entries) {
            $entries[$entry.FullName.Replace('\', '/')] = $entry
        }
        if (-not $entries.ContainsKey('recent-before.txt')) { throw 'Archive has no baseline manifest' }
        $reader = [IO.StreamReader]::new($entries['recent-before.txt'].Open())
        try { $manifest = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $verified = 0
        foreach ($line in ($manifest -split "`r?`n")) {
            if (-not $line -or $line -eq '<absent>' -or $line.StartsWith('directory|')) { continue }
            $parts = $line -split '\|'
            if ($parts.Count -ne 3) { throw 'Malformed baseline manifest in archive' }
            $name = 'original/Recent/' + $parts[0].Replace('\', '/')
            $entry = $entries[$name]
            if (-not $entry -or $entry.Length -ne [long]$parts[1]) {
                throw "Missing or truncated archived backup: $name"
            }
            $stream = $entry.Open()
            try {
                $hash = [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-', '')
            } finally { $stream.Dispose() }
            if ($hash -ne $parts[2]) { throw "Archived backup hash mismatch: $name" }
            $verified++
        }
        [ordered]@{ verified = $true; original_files_verified = $verified } |
            ConvertTo-Json |
            Set-Content -LiteralPath "$EvidenceDir\archive-verification.json" -Encoding UTF8
    } finally {
        $hasher.Dispose()
        $zip.Dispose()
    }
}

function Log([string]$Message) {
    Write-Host $Message
    # Start-Transcript includes the invoking command line and could leak -GuestPassword.
    "$(Get-Date -Format o) $Message" |
        Add-Content -LiteralPath "$EvidenceDir\host.log" -Encoding UTF8
}

try {
    Log "Build only on host; all tests run in VM: $Vmx"
    Push-Location $RepoRoot
    try {
        # Cargo's artifact message selects this exact build, not a stale glob match.
        $messages = & cargo test --release --test quick_access_manager --no-run --message-format=json
        if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed' }
        $artifacts = @($messages | ForEach-Object { $_ | ConvertFrom-Json } |
            Where-Object { $_.reason -eq 'compiler-artifact' -and
                $_.target.name -eq 'quick_access_manager' -and $_.executable })
        if ($artifacts.Count -ne 1) { throw 'Expected exactly one compiled test executable' }
        $exe = $artifacts[0].executable
        $commit = (& git rev-parse HEAD | Out-String).Trim()
        @(& rustc --version; & cargo --version) |
            Set-Content -LiteralPath "$EvidenceDir\toolchain.txt" -Encoding UTF8
        & git diff --no-ext-diff | Set-Content -LiteralPath "$EvidenceDir\source.diff" -Encoding UTF8
        & git status --short | Set-Content -LiteralPath "$EvidenceDir\source-status.txt" -Encoding UTF8
    } finally { Pop-Location }
    # Include untracked suite sources too; a commit ID alone cannot identify this build.
    New-Item -ItemType Directory -Path "$EvidenceDir\suite-source" | Out-Null
    Get-ChildItem -LiteralPath $PSScriptRoot -File -Force |
        Where-Object { $_.Name -ne '.env' } |
        Copy-Item -Destination "$EvidenceDir\suite-source"
    $hostHash = (Get-FileHash -LiteralPath $exe).Hash
    $running = (Vm @('list')).text -split "`r?`n"
    $wasRunning = @($running | Where-Object { $_.Trim() -ieq $Vmx }).Count -gt 0
    Log "Creating whole-VM rollback snapshot: $Snapshot"
    Vm @('snapshot', $Vmx, $Snapshot) | Out-Null
    $snapshotCreated = $true
    if (-not $wasRunning) { Vm @('start', $Vmx, 'nogui') | Out-Null }
    $deadline = (Get-Date).AddSeconds(240)
    do {
        $ready = Vm @('listProcessesInGuest', $Vmx) -Guest -AllowFailure
        if ($ready.code -eq 0) { break }
        if ((Get-Date) -ge $deadline) { throw 'Guest control unavailable after 240 seconds' }
        Start-Sleep -Seconds 3
    } while ($true)
    if ((Vm @('directoryExistsInGuest', $Vmx, 'C:\WincentTrack') -Guest -AllowFailure).code -ne 0) {
        Vm @('createDirectoryInGuest', $Vmx, 'C:\WincentTrack') -Guest | Out-Null
    }
    Vm @('createDirectoryInGuest', $Vmx, $GuestRoot) -Guest | Out-Null
    $config = [ordered]@{
        run_id = $RunId
        guest_user = $GuestUser
        platform = $Platform
        exe_sha256 = $hostHash
        source_commit = $commit
    }
    $config | ConvertTo-Json | Set-Content -LiteralPath "$EvidenceDir\config.json" -Encoding UTF8
    foreach ($pair in @(
        @($exe, "$GuestRoot\quick-access-manager.exe"),
        @("$PSScriptRoot\vm_quick_access_manager_guest.ps1", "$GuestRoot\guest.ps1"),
        @("$PSScriptRoot\vm_quick_access_manager_start.ps1", "$GuestRoot\start.ps1"),
        @("$EvidenceDir\config.json", "$GuestRoot\config.json")
    )) {
        Vm @('copyFileFromHostToGuest', $Vmx, $pair[0], $pair[1]) -Guest | Out-Null
    }
    Log 'Running comprehensive suite in the guest Explorer session'
    Vm @('runProgramInGuest', $Vmx, '-interactive',
        'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe',
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "$GuestRoot\start.ps1",
        '-ConfigPath', "$GuestRoot\config.json") -Guest | Out-Null
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    # FAILURE may appear before finally finishes. Always wait for finished.txt.
    $captured = $false
    while (-not (Guest-Exists "$GuestRoot\finished.txt")) {
        if (-not $captured -and (Guest-Exists "$GuestRoot\backup-ready.txt")) {
            Vm @('captureScreen', $Vmx, "$EvidenceDir\guest-running.png") `
                -Guest -AllowFailure | Out-Null
            $captured = $true
        }
        if ((Get-Date) -ge $deadline) { throw 'Guest completion timeout; snapshot rollback required' }
        Start-Sleep -Seconds 3
    }
    Pull-Evidence
    $result = Get-Content -LiteralPath "$EvidenceDir\result.json" -Raw | ConvertFrom-Json
    $final = Get-Content -LiteralPath "$EvidenceDir\final-state.json" -Raw | ConvertFrom-Json
    $environment = Get-Content -LiteralPath "$EvidenceDir\environment.json" -Raw | ConvertFrom-Json
    if ($result.run_id -ne $RunId -or $final.run_id -ne $RunId -or
        $environment.exe_sha256 -ne $hostHash -or $environment.platform -ne $Platform) {
        throw 'Evidence identity, platform or binary hash mismatch'
    }
    if (-not $result.passed -or -not $result.completed -or $result.checks -le 0 -or
        $result.failures -ne 0 -or -not $final.restored -or $final.restore_failure -or
        -not $final.targets_removed -or $final.remaining_test_processes -ne 0) {
        throw "Suite or restoration failed; inspect $EvidenceDir"
    }
    if (-not (Test-Path -LiteralPath "$EvidenceDir\evidence.zip") -or
        (Test-Path -LiteralPath "$EvidenceDir\ARTIFACT_FAILURE.txt")) {
        throw 'Evidence archive is incomplete'
    }
    Verify-Archive
    Log "Passed $($result.checks) checks; guest state restoration verified"
} catch {
    $failure = $_
    if ($snapshotCreated) {
        try { Pull-Evidence } catch { Write-Warning 'Could not retrieve every guest artifact' }
    }
} finally {
    if ($snapshotCreated) {
        try {
            Log 'Restoring the pre-test whole-VM snapshot'
            Vm @('revertToSnapshot', $Vmx, $Snapshot) | Out-Null
            if ($wasRunning) {
                $running = (Vm @('list')).text -split "`r?`n"
                if (@($running | Where-Object { $_.Trim() -ieq $Vmx }).Count -eq 0) {
                    Vm @('start', $Vmx, 'nogui') | Out-Null
                }
            }
            $running = (Vm @('list')).text -split "`r?`n"
            $nowRunning = @($running | Where-Object { $_.Trim() -ieq $Vmx }).Count -gt 0
            if ($nowRunning -ne $wasRunning) { throw 'VM power state differs from the baseline' }
            $rollback = $true
            # Retain failed runs' snapshot as a recovery aid.
            if ($null -eq $failure) { Vm @('deleteSnapshot', $Vmx, $Snapshot) | Out-Null }
        } catch {
            Write-Warning "Rollback/cleanup failed: $($_.Exception.Message)"
            if ($null -eq $failure) { $failure = $_ }
        }
    }
    [ordered]@{
        run_id = $RunId
        vmx = $Vmx
        snapshot = $Snapshot
        snapshot_created = $snapshotCreated
        snapshot_reverted = $rollback
        initially_running = $wasRunning
        passed = $null -eq $failure -and $rollback
        failure = if ($failure) { $failure.Exception.Message } else { $null }
    } | ConvertTo-Json | Set-Content -LiteralPath "$EvidenceDir\host-result.json" -Encoding UTF8
    Log "Host finished: snapshot_reverted=$rollback passed=$($null -eq $failure -and $rollback)"
}
if ($failure) { throw $failure }
Write-Host "Evidence: $EvidenceDir"
