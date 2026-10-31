# wincent get_item average-latency test - host orchestrator
#
# Runs on the Windows host. It builds the integration test in this directory,
# uploads the test binary plus the guest scripts to a dedicated Windows 10 or
# Windows 11 VM, starts the guest benchmark in Explorer's interactive session, waits
# for markers, captures screenshots, and pulls all evidence back.
#
# Isolation: every Quick Access mutation happens in the guest. This script only
# uses vmrun for file transfer, guest process control, marker polling and
# screenshots; it never calls wincent on the host.
#
# Configuration is read from `.env` next to this script (see `.env.example`).
# Explicit parameters override .env values. Both .env and command-line input are
# optional only for non-secret defaults; VMX/vmrun/guest credentials must be
# supplied somehow.
#
# Example:
#   pwsh -File tests/get_item_latency/run_get_item_latency_vm.ps1 `
#       -EvidenceDir temp/track_wincent/evidence/get-item-latency-20261027-v1

[CmdletBinding()]
param(
    [string]$Vmx = '',
    [string]$Vmrun = '',
    [string]$VmPassword = '',
    [string]$GuestUser = '',
    [string]$GuestPassword = '',
    [string]$EvidenceDir = '',
    [string]$Platform = '',
    [int]$Iters = 0,
    [int]$Warmup = -1,
    [string]$EnvFile = '',
    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $EnvFile) { $EnvFile = Join-Path $PSScriptRoot '.env' }
if ([System.IO.Path]::IsPathRooted($EnvFile) -eq $false) {
    $EnvFile = Join-Path $RepoRoot $EnvFile
}

$script:EnvValues = @{}
if (Test-Path -LiteralPath $EnvFile) {
    foreach ($line in Get-Content -LiteralPath $EnvFile) {
        $trimmed = $line.Trim()
        if (-not $trimmed -or $trimmed.StartsWith('#')) { continue }
        $separator = $trimmed.IndexOf('=')
        if ($separator -lt 1) { continue }
        $key = $trimmed.Substring(0, $separator).Trim()
        $value = $trimmed.Substring($separator + 1).Trim()
        if ($value.Length -ge 2) {
            $first = $value.Substring(0, 1)
            $last = $value.Substring($value.Length - 1, 1)
            if (($first -eq '"' -and $last -eq '"') -or ($first -eq "'" -and $last -eq "'")) {
                $value = $value.Substring(1, $value.Length - 2)
            }
        }
        $script:EnvValues[$key] = $value
    }
} elseif ($EnvFile -ne (Join-Path $PSScriptRoot '.env')) {
    throw "environment file not found: $EnvFile"
}

function Get-EnvSetting {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$Default = ''
    )
    if ($script:EnvValues.ContainsKey($Name) -and $script:EnvValues[$Name] -ne '') {
        return [string]$script:EnvValues[$Name]
    }
    return $Default
}

if (-not $Vmrun) { $Vmrun = Get-EnvSetting 'WINCENT_VM_VMRUN' }
if (-not $Vmx) { $Vmx = Get-EnvSetting 'WINCENT_VM_VMX' }
if (-not $VmPassword) { $VmPassword = Get-EnvSetting 'WINCENT_VM_PASSWORD' }
if (-not $GuestUser) { $GuestUser = Get-EnvSetting 'WINCENT_VM_GUEST_USER' }
if (-not $GuestPassword) { $GuestPassword = Get-EnvSetting 'WINCENT_VM_GUEST_PASSWORD' }
if (-not $Platform) { $Platform = Get-EnvSetting 'WINCENT_VM_PLATFORM' 'win10' }
$Platform = $Platform.ToLowerInvariant()
if ($Platform -notin @('win10', 'win11')) {
    throw "invalid platform '$Platform': expected win10 or win11"
}
if ($Iters -le 0) { $Iters = [int](Get-EnvSetting 'WINCENT_GET_ITEM_LATENCY_ITERS' '20') }
if ($Warmup -lt 0) { $Warmup = [int](Get-EnvSetting 'WINCENT_GET_ITEM_LATENCY_WARMUP' '2') }

$missing = @()
if (-not $Vmrun) { $missing += 'WINCENT_VM_VMRUN (or -Vmrun)' }
if (-not $Vmx) { $missing += 'WINCENT_VM_VMX (or -Vmx)' }
if (-not $GuestUser) { $missing += 'WINCENT_VM_GUEST_USER (or -GuestUser)' }
if (-not $GuestPassword) { $missing += 'WINCENT_VM_GUEST_PASSWORD (or -GuestPassword)' }
if ($missing.Count -gt 0) {
    throw "missing VM configuration: $($missing -join ', '). Copy .env.example to .env and fill it in."
}

$EvidenceRoot = Get-EnvSetting 'WINCENT_EVIDENCE_ROOT' (Join-Path $RepoRoot 'temp\track_wincent\evidence')
if (-not [System.IO.Path]::IsPathRooted($EvidenceRoot)) {
    $EvidenceRoot = Join-Path $RepoRoot $EvidenceRoot
}
if (-not $EvidenceDir) {
    $EvidenceDir = Join-Path $EvidenceRoot ("get-item-latency-$Platform-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
} elseif (-not [System.IO.Path]::IsPathRooted($EvidenceDir)) {
    $EvidenceDir = Join-Path $RepoRoot $EvidenceDir
}
$EvidenceDir = [System.IO.Path]::GetFullPath($EvidenceDir)

$GuestRoot = 'C:\WincentTrack\get-item-latency-current'
$GuestExe = 'C:\WincentTrack\wincent-get-item-latency.exe'
$GuestScript = 'C:\WincentTrack\run-get-item-latency.ps1'
$GuestStart = 'C:\WincentTrack\start-get-item-latency.ps1'
$GuestConfig = 'C:\WincentTrack\get-item-latency-config.json'
$GuestArchive = 'C:\WincentTrack\get-item-latency-evidence.zip'
$TaskName = 'WincentTrack-GetItemLatency'
$RunId = Split-Path -Leaf $EvidenceDir
$GuestFailureFiles = @("$GuestRoot\FAILURE.txt", "$GuestRoot\RESTORE_FAILURE.txt")

function Invoke-Vmrun {
    param(
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [switch]$AllowFailure
    )
    $vmArgs = @('-T', 'ws')
    if ($VmPassword) { $vmArgs += @('-vp', $VmPassword) }
    $output = & $Vmrun @vmArgs @Arguments 2>&1
    $code = $LASTEXITCODE
    $text = ($output | Out-String).Trim()
    if ($code -ne 0 -and -not $AllowFailure) {
        throw "vmrun failed ($code): $($Arguments -join ' ')`n$text"
    }
    return [pscustomobject]@{ ExitCode = $code; Output = $text }
}

function Invoke-GuestVmrun {
    param(
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [switch]$AllowFailure
    )
    $guestArgs = @('-gu', $GuestUser, '-gp', $GuestPassword) + $Arguments
    return Invoke-Vmrun -Arguments $guestArgs -AllowFailure:$AllowFailure
}

function Test-GuestPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    $result = Invoke-GuestVmrun -Arguments @('fileExistsInGuest', $Vmx, $Path) -AllowFailure
    return $result.ExitCode -eq 0
}

function Assert-NoGuestFailure {
    foreach ($failureFile in $GuestFailureFiles) {
        if (Test-GuestPath $failureFile) {
            throw "guest reported failure: $failureFile"
        }
    }
}

function Wait-GuestPath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$TimeoutSeconds = 1200,
        [string]$Label = $Path,
        [switch]$IgnoreFailureMarkers
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (-not $IgnoreFailureMarkers) { Assert-NoGuestFailure }
        if (Test-GuestPath $Path) { return }
        Start-Sleep -Seconds 2
    }
    throw "timed out after ${TimeoutSeconds}s waiting for $Label"
}

function Copy-ToGuest {
    param([string]$HostPath, [string]$GuestPath)
    Invoke-GuestVmrun -Arguments @('copyFileFromHostToGuest', $Vmx, $HostPath, $GuestPath) | Out-Null
}

function Copy-FromGuest {
    param([string]$GuestPath, [string]$HostPath)
    Invoke-GuestVmrun -Arguments @('copyFileFromGuestToHost', $Vmx, $GuestPath, $HostPath) | Out-Null
}

function Capture-GuestScreen {
    param([string]$HostPath)
    Invoke-GuestVmrun -Arguments @('captureScreen', $Vmx, $HostPath) | Out-Null
}

function Pull-Artifacts {
    $names = @(
        'result.json', 'timings.csv', 'final-state.json', 'environment.json',
        'benchmark.log', 'FAILURE.txt', 'RESTORE_FAILURE.txt', 'ARCHIVE_FAILURE.txt'
    )
    foreach ($name in $names) {
        $guestPath = "$GuestRoot\$name"
        if (Test-GuestPath $guestPath) {
            Copy-FromGuest $guestPath (Join-Path $EvidenceDir $name)
        }
    }
    if (Test-GuestPath $GuestArchive) {
        $zip = Join-Path $EvidenceDir 'get-item-latency-evidence.zip'
        Copy-FromGuest $GuestArchive $zip
        $extracted = Join-Path $EvidenceDir 'extracted'
        if (Test-Path -LiteralPath $extracted) { Remove-Item -LiteralPath $extracted -Recurse -Force }
        Expand-Archive -LiteralPath $zip -DestinationPath $extracted -Force
    }
}

function Remove-GuestTask {
    try {
        Invoke-GuestVmrun -Arguments @(
            'runProgramInGuest', $Vmx, 'C:\Windows\System32\schtasks.exe',
            '/Delete', '/TN', $TaskName, '/F'
        ) -AllowFailure | Out-Null
    } catch {
        Write-Warning "failed to delete guest scheduled task: $($_.Exception.Message)"
    }
}

New-Item -ItemType Directory -Path $EvidenceDir -Force | Out-Null
Start-Transcript -Path (Join-Path $EvidenceDir 'host-runner.log') -Force | Out-Null

$failure = $null
try {
    Write-Host "evidence: $EvidenceDir"
    if (-not (Test-Path -LiteralPath $Vmrun)) { throw "vmrun not found: $Vmrun" }
    if (-not (Test-Path -LiteralPath $Vmx)) { throw "VMX not found: $Vmx" }

    $list = Invoke-Vmrun -Arguments @('list')
    if ($list.Output -notmatch [regex]::Escape($Vmx)) {
        Write-Host "starting VM: $Vmx"
        Invoke-Vmrun -Arguments @('start', $Vmx, 'nogui') | Out-Null
    }
    $ip = $null
    $ipResult = Invoke-Vmrun -Arguments @('getGuestIPAddress', $Vmx, '-wait') -AllowFailure
    if ($ipResult.ExitCode -eq 0) { $ip = $ipResult.Output } else { Write-Warning 'getGuestIPAddress timed out; continuing with guest control' }
    $deadline = (Get-Date).AddSeconds(240)
    while ($true) {
        $probe = Invoke-GuestVmrun -Arguments @('listProcessesInGuest', $Vmx) -AllowFailure
        if ($probe.ExitCode -eq 0) { break }
        if ((Get-Date) -gt $deadline) { throw 'guest control is not available' }
        Start-Sleep -Seconds 3
    }
    Write-Host "guest control ready; ip=$(if ($ip) { $ip } else { 'unknown' })"

    if (-not $SkipBuild) {
        Write-Host 'building tests/get_item_latency/main.rs'
        Push-Location $RepoRoot
        try {
            & cargo test --release --test get_item_latency --no-run
            if ($LASTEXITCODE -ne 0) { throw "cargo test --no-run failed with exit $LASTEXITCODE" }
        } finally {
            Pop-Location
        }
    }
    $exe = Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'target\release\deps') -Filter 'get_item_latency-*.exe' -File |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $exe) { throw 'get_item_latency test binary was not found' }
    $hostHash = (Get-FileHash -LiteralPath $exe.FullName -Algorithm SHA256).Hash
    Write-Host "test binary: $($exe.FullName) sha256=$hostHash"

    $guestDirectory = Invoke-GuestVmrun -Arguments @(
        'directoryExistsInGuest', $Vmx, 'C:\WincentTrack'
    ) -AllowFailure
    if ($guestDirectory.ExitCode -ne 0) {
        Invoke-GuestVmrun -Arguments @(
            'createDirectoryInGuest', $Vmx, 'C:\WincentTrack'
        ) | Out-Null
    }
    Copy-ToGuest $exe.FullName $GuestExe
    Copy-ToGuest (Join-Path $PSScriptRoot 'vm_get_item_latency_guest.ps1') $GuestScript
    Copy-ToGuest (Join-Path $PSScriptRoot 'vm_get_item_latency_start.ps1') $GuestStart
    $config = [ordered]@{
        run_id = $RunId
        exe = $GuestExe
        guest_user = $GuestUser
        platform = $Platform
        iters = $Iters
        warmup = $Warmup
    }
    $configPath = Join-Path $EvidenceDir 'config.json'
    $config | ConvertTo-Json | Set-Content -LiteralPath $configPath -Encoding UTF8
    Copy-ToGuest $configPath $GuestConfig

    Write-Host "launching guest benchmark (platform=$Platform iters=$Iters warmup=$Warmup)"
    Invoke-GuestVmrun -Arguments @(
        'runProgramInGuest', $Vmx,
        '-interactive',
        'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe',
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $GuestStart
    ) | Out-Null

    # The launcher deletes the previous run's root and failure markers before it
    # writes this run-specific marker, so stale failure files cannot race with
    # the start of this run.
    Wait-GuestPath -Path "C:\WincentTrack\get-item-latency-started-$RunId.txt" `
        -TimeoutSeconds 180 -Label 'guest run started' -IgnoreFailureMarkers

    Wait-GuestPath -Path "$GuestRoot\markers\10-baseline-ready.txt" -Label '10-baseline-ready'
    Capture-GuestScreen (Join-Path $EvidenceDir '10-baseline-ready.png')
    Wait-GuestPath -Path "$GuestRoot\markers\80-benchmark-complete.txt" -Label '80-benchmark-complete' -TimeoutSeconds 2400
    Wait-GuestPath -Path "$GuestRoot\markers\90-restored.txt" -Label '90-restored' -TimeoutSeconds 600
    Capture-GuestScreen (Join-Path $EvidenceDir '90-restored.png')
    if (-not (Test-GuestPath "$GuestRoot\final-state.json")) {
        throw 'final-state.json was not written'
    }

    Pull-Artifacts

    $finalState = Get-Content -LiteralPath (Join-Path $EvidenceDir 'final-state.json') -Raw | ConvertFrom-Json
    if (-not $finalState.restored) { throw 'guest state was not restored to the baseline' }
    $environment = Get-Content -LiteralPath (Join-Path $EvidenceDir 'environment.json') -Raw | ConvertFrom-Json
    if ($environment.exe_sha256 -ne $hostHash) {
        throw "guest executable hash $($environment.exe_sha256) != host $hostHash"
    }
    if ($environment.detected_platform -ne $Platform) {
        throw "guest platform $($environment.detected_platform) != requested $Platform"
    }
    foreach ($failureName in @('FAILURE.txt', 'RESTORE_FAILURE.txt')) {
        if (Test-Path -LiteralPath (Join-Path $EvidenceDir $failureName)) {
            throw "guest failure marker present: $failureName"
        }
    }

    $result = Get-Content -LiteralPath (Join-Path $EvidenceDir 'result.json') -Raw | ConvertFrom-Json
    if ($result.platform -ne $Platform) { throw "result platform $($result.platform) != requested $Platform" }
    $requiredScenarios = if ($Platform -eq 'win10') {
        @('recent_full', 'frequent_full', 'recent_under', 'frequent_under')
    } else {
        @('frequent_native')
    }
    $requiredScenarios = @($requiredScenarios)
    if ($Platform -eq 'win11' -and $result.fixture_capabilities.recent) {
        $requiredScenarios += 'recent_native'
    }
    foreach ($requiredScenario in $requiredScenarios) {
        if (-not $result.scenarios.PSObject.Properties[$requiredScenario]) {
            throw "result.json is missing scenario $requiredScenario"
        }
        foreach ($requiredMethod in @('shell_query', 'cfb_parse')) {
            if (-not $result.scenarios.$requiredScenario.PSObject.Properties[$requiredMethod]) {
                throw "result.json is missing $requiredScenario/$requiredMethod"
            }
        }
    }
    Write-Host ''
    Write-Host ('{0,-16} {1,-12} {2,6} {3,12} {4,12} {5,12} {6,12}' -f 'scenario', 'method', 'n', 'mean_ms', 'median_ms', 'p95_ms', 'max_ms')
    foreach ($scenarioProperty in $result.scenarios.PSObject.Properties) {
        foreach ($methodName in @('shell_query', 'cfb_parse')) {
            $stats = $scenarioProperty.Value.$methodName
            if (-not $stats) { continue }
            Write-Host ('{0,-16} {1,-12} {2,6} {3,12:N3} {4,12:N3} {5,12:N3} {6,12:N3}' -f `
                $scenarioProperty.Name, $methodName, $stats.count, $stats.mean_ms, $stats.median_ms, $stats.p95_ms, $stats.max_ms)
        }
    }
    Write-Host ''
    Write-Host "done: $EvidenceDir"
} catch {
    $failure = $_
    try { Pull-Artifacts } catch { Write-Warning "failed to pull artifacts: $($_.Exception.Message)" }
    throw
} finally {
    Remove-GuestTask
    try { Stop-Transcript | Out-Null } catch { }
}
