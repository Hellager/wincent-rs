# wincent get_item average-latency test - guest orchestrator (Windows PowerShell 5.1)
#
# Runs inside the dedicated Windows test VM in the interactive Explorer session
# (session 1). It is launched by start-get-item-latency.ps1 through a temporary
# scheduled task so that the benchmark shares Explorer's session and COM state.
#
# Flow:
#   1. verify session and locate both Quick Access backing files;
#   2. back up Recent/Frequent CFB files plus the complete Recent .lnk set;
#   3. run the Rust integration-test binary (tests/get_item_latency/main.rs);
#      Windows 10 builds deterministic v4 fixtures, while Windows 11 creates
#      native v6 state through system APIs and records available capabilities;
#   4. stop Explorer, restore both CFB files and the Recent .lnk set, restart
#      Explorer, and verify SHA-256/link-manifest equality;
#   5. write markers, environment.json, final-state.json and a zip archive.
#
# All mutations stay in the guest. The host only copies files and polls markers.

param(
    [string]$Root = 'C:\WincentTrack\get-item-latency-current',
    [string]$ConfigPath = 'C:\WincentTrack\get-item-latency-config.json'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$RecentDir = Join-Path $env:APPDATA 'Microsoft\Windows\Recent'
$AutoDir = Join-Path $RecentDir 'AutomaticDestinations'
$RecentCfb = Join-Path $AutoDir '5f7b5f1e01b83767.automaticDestinations-ms'
$FrequentCfb = Join-Path $AutoDir 'f01b4d95cf55d32a.automaticDestinations-ms'
$TargetsDir = Join-Path $Root 'targets'
$OriginalDir = Join-Path $Root 'original'
$MarkersDir = Join-Path $Root 'markers'
$ResultPath = Join-Path $Root 'result.json'
$RowsPath = Join-Path $Root 'timings.csv'
$LogPath = Join-Path $Root 'benchmark.log'
$FailurePath = Join-Path $Root 'FAILURE.txt'
$RestoreFailurePath = Join-Path $Root 'RESTORE_FAILURE.txt'
$FinalStatePath = Join-Path $Root 'final-state.json'
$EnvironmentPath = Join-Path $Root 'environment.json'
$ArchivePath = 'C:\WincentTrack\get-item-latency-evidence.zip'
$QuickAccessNamespace = 'shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}'

$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$Exe = [string]$config.exe
$Iters = [int]$config.iters
$Warmup = [int]$config.warmup
$RunId = [string]$config.run_id
$RequestedPlatform = ([string]$config.platform).ToLowerInvariant()
if ($RequestedPlatform -notin @('win10', 'win11')) {
    throw "config platform must be win10 or win11, got '$RequestedPlatform'"
}

function Get-LinkManifestText([string] $Directory) {
    return ((Get-ChildItem -LiteralPath $Directory -Filter '*.lnk' -File |
        Sort-Object Name | ForEach-Object {
            "$($_.Name)|$($_.Length)|$((Get-FileHash $_.FullName -Algorithm SHA256).Hash)"
        }) -join "`n")
}

function Get-BackingState([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{ state = 'absent'; length = 0; sha256 = $null }
    }
    $item = Get-Item -LiteralPath $Path
    return [pscustomobject]@{
        state = if ($item.Length -eq 0) { 'zero_length' } else { 'populated' }
        length = $item.Length
        sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    }
}

function Save-BackingFile([string] $Source, [string] $Destination) {
    $state = Get-BackingState $Source
    if ($state.state -ne 'absent') {
        Copy-Item -LiteralPath $Source -Destination $Destination -Force
    }
    return $state
}

function Restore-BackingFile($State, [string] $Backup, [string] $Destination) {
    if ($State.state -eq 'absent') {
        Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
        return
    }
    Copy-Item -LiteralPath $Backup -Destination $Destination -Force
}

function Test-BackingRestored($Expected, [string] $Path) {
    $actual = Get-BackingState $Path
    return $actual.state -eq $Expected.state -and
        [int64]$actual.length -eq [int64]$Expected.length -and
        $actual.sha256 -eq $Expected.sha256
}

function Save-State([string] $Destination) {
    $links = Join-Path $Destination 'links'
    New-Item -ItemType Directory -Path $Destination, $links -Force | Out-Null
    $backingState = [ordered]@{
        recent = Save-BackingFile $RecentCfb (Join-Path $Destination 'recent.cfb')
        frequent = Save-BackingFile $FrequentCfb (Join-Path $Destination 'frequent.cfb')
    }
    Get-ChildItem -LiteralPath $RecentDir -Filter '*.lnk' -File |
        Copy-Item -Destination $links -Force
    Get-LinkManifestText $RecentDir |
        Set-Content -LiteralPath (Join-Path $Destination 'links-manifest.txt') -Encoding UTF8
    $backingState | ConvertTo-Json -Depth 4 |
        Set-Content -LiteralPath (Join-Path $Destination 'backing-state.json') -Encoding UTF8
}

function Stop-ExplorerSession {
    Get-Process explorer -ErrorAction SilentlyContinue |
        Where-Object SessionId -eq $script:ScriptSession |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 1200
}

function Start-ExplorerSession {
    Start-Process explorer.exe $QuickAccessNamespace
    Start-Sleep -Milliseconds 2500
}

function Restore-State {
    Stop-ExplorerSession
    $state = Get-Content -LiteralPath (Join-Path $OriginalDir 'backing-state.json') -Raw |
        ConvertFrom-Json
    Restore-BackingFile $state.recent (Join-Path $OriginalDir 'recent.cfb') $RecentCfb
    Restore-BackingFile $state.frequent (Join-Path $OriginalDir 'frequent.cfb') $FrequentCfb
    Get-ChildItem -LiteralPath $RecentDir -Filter '*.lnk' -File | Remove-Item -Force
    # Remove a stray "links" directory if an earlier run copied the backup folder
    # itself into the Recent directory instead of its children.
    $stray = Join-Path $RecentDir 'links'
    if (Test-Path -LiteralPath $stray) {
        Remove-Item -LiteralPath $stray -Recurse -Force
    }
    Get-ChildItem -LiteralPath (Join-Path $OriginalDir 'links') -Filter '*.lnk' -File |
        Copy-Item -Destination $RecentDir -Force
    Start-ExplorerSession
}

if (-not (Test-Path -LiteralPath $Exe)) { throw "benchmark executable is missing: $Exe" }
New-Item -ItemType Directory -Path $AutoDir -Force | Out-Null

$script:ScriptSession = (Get-Process -Id $PID).SessionId
$initialExplorer = Get-Process explorer -ErrorAction Stop |
    Where-Object SessionId -eq $script:ScriptSession |
    Select-Object -First 1
if (-not $initialExplorer) {
    throw "No Explorer process in script session $script:ScriptSession"
}
if ($script:ScriptSession -ne 1) {
    throw "Benchmark must run in interactive session 1, got $script:ScriptSession"
}

$os = Get-CimInstance Win32_OperatingSystem
$DetectedPlatform = if ([int]$os.BuildNumber -ge 22000) { 'win11' } else { 'win10' }
if ($DetectedPlatform -ne $RequestedPlatform) {
    throw "guest platform $DetectedPlatform (build $($os.BuildNumber)) does not match requested $RequestedPlatform"
}

if (Test-Path -LiteralPath $Root) { Remove-Item -LiteralPath $Root -Recurse -Force }
New-Item -ItemType Directory -Path $Root, $TargetsDir, $MarkersDir -Force | Out-Null

$failure = $null
$restored = $false

try {
    Save-State $OriginalDir

    $backingState = Get-Content -LiteralPath (Join-Path $OriginalDir 'backing-state.json') -Raw |
        ConvertFrom-Json
    [pscustomobject]@{
        run_id = $RunId
        timestamp = (Get-Date).ToString('o')
        computer_name = $env:COMPUTERNAME
        user = $env:USERNAME
        os_caption = $os.Caption
        os_version = $os.Version
        os_build = $os.BuildNumber
        requested_platform = $RequestedPlatform
        detected_platform = $DetectedPlatform
        powershell_version = $PSVersionTable.PSVersion.ToString()
        script_pid = $PID
        script_session = $script:ScriptSession
        explorer_pid = $initialExplorer.Id
        explorer_session = $initialExplorer.SessionId
        exe = $Exe
        exe_sha256 = (Get-FileHash -LiteralPath $Exe -Algorithm SHA256).Hash
        recent_cfb = $RecentCfb
        frequent_cfb = $FrequentCfb
        recent_backing_state = $backingState.recent.state
        frequent_backing_state = $backingState.frequent.state
        recent_sha256 = $backingState.recent.sha256
        frequent_sha256 = $backingState.frequent.sha256
        iterations = $Iters
        warmup = $Warmup
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $EnvironmentPath -Encoding UTF8

    (Get-Date).ToString('o') |
        Set-Content -LiteralPath (Join-Path $MarkersDir '10-baseline-ready.txt') -Encoding UTF8

    $env:WINCENT_GET_ITEM_LATENCY = '1'
    $env:WINCENT_TEST_PLATFORM = $DetectedPlatform
    $env:WINCENT_GET_ITEM_LATENCY_OUT = $ResultPath
    $env:WINCENT_GET_ITEM_LATENCY_ROOT = $TargetsDir
    $env:WINCENT_GET_ITEM_LATENCY_ITERS = [string]$Iters
    $env:WINCENT_GET_ITEM_LATENCY_WARMUP = [string]$Warmup

    # Run through a cmd wrapper: cmd waits for the test process and exits even
    # if a long-lived child such as Explorer inherited a standard handle.
    $wrapperPath = Join-Path $Root 'run-benchmark.cmd'
    $wrapper = @(
        '@echo off',
        "`"$Exe`" get_item_latency --ignored --exact --nocapture > `"$LogPath`" 2>&1",
        'exit /b %ERRORLEVEL%'
    ) -join "`r`n"
    Set-Content -LiteralPath $wrapperPath -Value $wrapper -Encoding ASCII
    & cmd.exe /c $wrapperPath
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        $tail = Get-Content -LiteralPath $LogPath -Tail 40 -ErrorAction SilentlyContinue
        throw "benchmark exited with $exitCode`n$($tail -join "`n")"
    }
    if (-not (Test-Path -LiteralPath $ResultPath)) {
        throw "benchmark did not write $ResultPath"
    }

    $result = Get-Content -LiteralPath $ResultPath -Raw | ConvertFrom-Json
    if ([int]$result.iterations -ne $Iters) {
        throw "result iterations $($result.iterations) != requested $Iters"
    }
    $requiredScenarios = if ($DetectedPlatform -eq 'win10') {
        @('recent_full', 'frequent_full', 'recent_under', 'frequent_under')
    } else {
        @('frequent_native')
    }
    $requiredScenarios = @($requiredScenarios)
    if ($DetectedPlatform -eq 'win11' -and $result.fixture_capabilities.recent) {
        $requiredScenarios += 'recent_native'
    }
    foreach ($scenario in $requiredScenarios) {
        if (-not $result.scenarios.PSObject.Properties[$scenario]) {
            throw "result is missing scenario $scenario"
        }
    }
    if (-not (Test-Path -LiteralPath $RowsPath)) {
        throw "benchmark did not write $RowsPath"
    }

    (Get-Date).ToString('o') |
        Set-Content -LiteralPath (Join-Path $MarkersDir '80-benchmark-complete.txt') -Encoding UTF8
} catch {
    $failure = $_ | Out-String
    $failure | Set-Content -LiteralPath $FailurePath -Encoding UTF8
} finally {
    try {
        Restore-State
        $state = Get-Content -LiteralPath (Join-Path $OriginalDir 'backing-state.json') -Raw |
            ConvertFrom-Json
        $links = (Get-Content -LiteralPath (Join-Path $OriginalDir 'links-manifest.txt') -Raw).TrimEnd()
        $restored =
            (Test-BackingRestored $state.recent $RecentCfb) -and
            (Test-BackingRestored $state.frequent $FrequentCfb) -and
            ((Get-LinkManifestText $RecentDir) -eq $links)
        if (-not $restored) { throw 'Final CFB hashes or Recent link manifest do not match baseline' }
    } catch {
        $_ | Out-String | Set-Content -LiteralPath $RestoreFailurePath -Encoding UTF8
    }

    Remove-Item -LiteralPath $TargetsDir -Recurse -Force -ErrorAction SilentlyContinue

    [pscustomobject]@{
        run_id = $RunId
        restored = $restored
        failure = $failure
        recent_sha256 = if (Test-Path $RecentCfb) { (Get-FileHash $RecentCfb -Algorithm SHA256).Hash } else { $null }
        frequent_sha256 = if (Test-Path $FrequentCfb) { (Get-FileHash $FrequentCfb -Algorithm SHA256).Hash } else { $null }
        explorer_session = @(Get-Process explorer -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty SessionId)
        timestamp = (Get-Date).ToString('o')
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $FinalStatePath -Encoding UTF8

    Remove-Item -LiteralPath $ArchivePath -Force -ErrorAction SilentlyContinue
    try {
        Compress-Archive -LiteralPath $Root -DestinationPath $ArchivePath -Force
    } catch {
        $_ | Out-String | Set-Content -LiteralPath (Join-Path $Root 'ARCHIVE_FAILURE.txt') -Encoding UTF8
    }

    (Get-Date).ToString('o') |
        Set-Content -LiteralPath (Join-Path $MarkersDir '90-restored.txt') -Encoding UTF8
}

if ($failure) { throw $failure }
