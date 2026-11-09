# Windows PowerShell 5.1. Run only through the VM host runner.
param([Parameter(Mandatory = $true)][string]$ConfigPath)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
if ([string]$config.run_id -notmatch '^[a-zA-Z0-9-]+$') { throw 'Invalid run ID' }
$Root = "C:\WincentTrack\qam-$($config.run_id)"
if ([System.IO.Path]::GetFullPath($ConfigPath) -ne "$Root\config.json") {
    throw 'Config must reside in its dedicated VM run directory'
}
$Exe = "$Root\quick-access-manager.exe"
$UnitExe = "$Root\standalone-unit-tests.exe"
$Recent = Join-Path $env:APPDATA 'Microsoft\Windows\Recent'
$Backup = Join-Path $Root 'original'
$RegistrySpecs = @(
    @('Software\Microsoft\Windows\CurrentVersion\Explorer', 'ShowRecent'),
    @('Software\Microsoft\Windows\CurrentVersion\Explorer', 'ShowFrequent'),
    @('Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced', 'Start_TrackDocs')
)
$Session = (Get-Process -Id $PID).SessionId
$failure = $null
$restoreFailure = $null
$backupReady = $false
$restored = $false
$exitCode = $null
$process = $null
$recentExisted = $false
$registry = @()
$script:ExplorerTouched = $false

function Write-Json($Value, [string]$Name) {
    $Value | ConvertTo-Json -Depth 12 |
        Set-Content -LiteralPath (Join-Path $Root $Name) -Encoding UTF8
}

function Get-Manifest([string]$Directory) {
    if (-not (Test-Path -LiteralPath $Directory)) { return '<absent>' }
    $prefix = [System.IO.Path]::GetFullPath($Directory).TrimEnd('\') + '\'
    $lines = @(Get-ChildItem -LiteralPath $Directory -Recurse -Force |
        Sort-Object FullName | ForEach-Object {
            if ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Refusing reparse point in Recent backup: $($_.FullName)"
            }
            $relative = $_.FullName.Substring($prefix.Length)
            if ($_.PSIsContainer) { "directory|$relative" } else {
                "$relative|$($_.Length)|$((Get-FileHash -LiteralPath $_.FullName).Hash)"
            }
        })
    return $lines -join "`n"
}

function Get-RegistryState {
    foreach ($spec in $RegistrySpecs) {
        $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($spec[0])
        try {
            $exists = $null -ne $key -and $key.GetValueNames() -contains $spec[1]
            [pscustomobject]@{
                key = $spec[0]
                name = $spec[1]
                exists = $exists
                kind = if ($exists) { $key.GetValueKind($spec[1]).ToString() } else { $null }
                value = if ($exists) { $key.GetValue($spec[1]) } else { $null }
            }
        } finally {
            if ($null -ne $key) { $key.Dispose() }
        }
    }
}

function Stop-Explorer {
    $script:ExplorerTouched = $true
    Get-Process explorer -ErrorAction SilentlyContinue |
        Where-Object SessionId -eq $Session | Stop-Process -Force
    Start-Sleep -Seconds 2
}

function Start-Explorer {
    Start-Process -FilePath "$env:WINDIR\explorer.exe" -WindowStyle Hidden
    Start-Sleep -Seconds 3
}

function Clear-RecentChildren {
    $expected = [IO.Path]::GetFullPath((Join-Path $env:APPDATA 'Microsoft\Windows\Recent'))
    if ([IO.Path]::GetFullPath($Recent) -ne $expected) { throw 'Recent path mismatch' }
    if (Test-Path -LiteralPath $Recent) {
        # Get-Manifest rejects junctions/symlinks before this recursive removal.
        $null = Get-Manifest $Recent
        Get-ChildItem -LiteralPath $Recent -Force | ForEach-Object {
            $resolved = [IO.Path]::GetFullPath($_.FullName)
            if (-not $resolved.StartsWith($expected + '\', [StringComparison]::OrdinalIgnoreCase)) {
                throw "Unsafe restore path: $resolved"
            }
            Remove-Item -LiteralPath $resolved -Recurse -Force
        }
    }
}

try {
    $system = Get-CimInstance Win32_ComputerSystem
    if ($system.Manufacturer -notmatch 'VMware' -or $system.Model -notmatch 'VMware') {
        throw 'This destructive suite requires a VMware guest'
    }
    $explorer = Get-Process explorer -ErrorAction Stop |
        Where-Object SessionId -eq $Session | Select-Object -First 1
    if ($Session -eq 0 -or -not $explorer) { throw 'Runner must share the interactive Explorer session' }
    if ($env:USERNAME -ine [string]$config.guest_user.Split('\')[-1]) {
        throw 'Guest user does not match the requested test profile'
    }
    $os = Get-CimInstance Win32_OperatingSystem
    $platform = if ([int]$os.BuildNumber -ge 22000) { 'win11' } else { 'win10' }
    if ($platform -ne $config.platform) { throw "Platform mismatch: $platform != $($config.platform)" }
    $exeHash = (Get-FileHash -LiteralPath $Exe).Hash
    if ($exeHash -ne $config.exe_sha256) { throw 'Uploaded executable hash mismatch' }
    $unitHash = (Get-FileHash -LiteralPath $UnitExe).Hash
    if ($unitHash -ne $config.unit_exe_sha256) { throw 'Uploaded unit executable hash mismatch' }
    $version = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $dlls = @('shell32.dll', 'windows.storage.dll', 'explorer.exe') | ForEach-Object {
        $path = if ($_ -eq 'explorer.exe') { "$env:WINDIR\$_" } else { "$env:WINDIR\System32\$_" }
        [pscustomobject]@{
            path = $path
            version = (Get-Item -LiteralPath $path).VersionInfo.FileVersion
            sha256 = (Get-FileHash -LiteralPath $path).Hash
        }
    }
    Write-Json ([ordered]@{
        run_id = $config.run_id
        timestamp = (Get-Date).ToString('o')
        computer = $env:COMPUTERNAME
        user = $env:USERNAME
        os_caption = $os.Caption
        os_build = $os.BuildNumber
        ubr = $version.UBR
        display_version = $version.DisplayVersion
        platform = $platform
        script_session = $Session
        explorer_session = $explorer.SessionId
        exe_sha256 = $exeHash
        unit_exe_sha256 = $unitHash
        source_commit = $config.source_commit
        dlls = @($dlls)
    }) 'environment.json'

    # Stop Explorer during backup to avoid combining different persistence generations.
    $registry = @(Get-RegistryState)
    Write-Json $registry 'registry-before.json'
    Stop-Explorer
    New-Item -ItemType Directory -Path $Backup -ErrorAction Stop | Out-Null
    $recentExisted = Test-Path -LiteralPath $Recent
    $before = Get-Manifest $Recent
    if ($recentExisted) {
        New-Item -ItemType Directory -Path "$Backup\Recent" | Out-Null
        Get-ChildItem -LiteralPath $Recent -Force |
            Copy-Item -Destination "$Backup\Recent" -Recurse -Force
        if ((Get-Manifest "$Backup\Recent") -cne $before) { throw 'Recent backup is inconsistent' }
    }
    $before | Set-Content -LiteralPath "$Root\recent-before.txt" -Encoding UTF8
    $backupReady = $true
    Set-Content -LiteralPath "$Root\backup-ready.txt" -Value $config.run_id
    Start-Explorer
    Start-Process -FilePath "$env:WINDIR\explorer.exe" -WindowStyle Hidden `
        -ArgumentList 'shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}'
    Start-Sleep -Seconds 2

    $env:WINCENT_QAM_ROOT = $Root
    $env:WINCENT_QAM_RUN_ID = [string]$config.run_id
    $env:WINCENT_QAM_VM_ACK = [string]$config.run_id
    $env:WINCENT_QAM_GUEST = $env:COMPUTERNAME
    $wrapper = @(
        '@echo off',
        "`"$UnitExe`" standalone::tests:: --test-threads=1 --nocapture > `"$Root\unit-tests.log`" 2>&1",
        'if errorlevel 1 exit /b %ERRORLEVEL%',
        "`"$Exe`" --ignored --exact quick_access_manager --nocapture > `"$Root\test.log`" 2>&1",
        'exit /b %ERRORLEVEL%'
    ) -join "`r`n"
    Set-Content -LiteralPath "$Root\run-test.cmd" -Value $wrapper -Encoding ASCII
    $process = Start-Process -FilePath "$env:WINDIR\System32\cmd.exe" `
        -ArgumentList @('/c', "$Root\run-test.cmd") -PassThru -WindowStyle Hidden
    if (-not $process.WaitForExit(900000)) { throw 'Test executable exceeded 900 seconds' }
    $process.Refresh()
    $exitCode = $process.ExitCode
    if ($exitCode -ne 0) { throw "Test executable failed: exit=$exitCode (see test.log)" }
} catch {
    $failure = $_ | Out-String
    $failure | Set-Content -LiteralPath "$Root\FAILURE.txt" -Encoding UTF8
} finally {
    # A timeout never cancels an in-flight COM call; terminate this run's process
    # before restoring, and use the host snapshot as the final isolation boundary.
    try {
        Get-CimInstance Win32_Process `
            -Filter "Name='quick-access-manager.exe' OR Name='standalone-unit-tests.exe'" |
            Where-Object { $_.ExecutablePath -ieq $Exe -or $_.ExecutablePath -ieq $UnitExe } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop }
        if ($null -ne $process -and -not $process.HasExited) { $process.Kill(); $process.WaitForExit() }
        if ($backupReady) {
            Stop-Explorer
            foreach ($entry in $registry) {
                $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($entry.key)
                try {
                    if ($entry.exists) {
                        $kind = [Enum]::Parse([Microsoft.Win32.RegistryValueKind], $entry.kind)
                        $key.SetValue($entry.name, $entry.value, $kind)
                    } else {
                        $key.DeleteValue($entry.name, $false)
                    }
                } finally { $key.Dispose() }
            }
            Clear-RecentChildren
            if ($recentExisted) {
                New-Item -ItemType Directory -Path $Recent -Force | Out-Null
                Get-ChildItem -LiteralPath "$Backup\Recent" -Force |
                    Copy-Item -Destination $Recent -Recurse -Force
            } elseif (Test-Path -LiteralPath $Recent) {
                Remove-Item -LiteralPath $Recent -Force
            }
            $after = Get-Manifest $Recent
            $after | Set-Content -LiteralPath "$Root\recent-restored.txt" -Encoding UTF8
            $registryAfter = @(Get-RegistryState)
            Write-Json $registryAfter 'registry-restored.json'
            $registryRestored = ($registry | ConvertTo-Json -Compress) -ceq `
                ($registryAfter | ConvertTo-Json -Compress)
            $restored = $after -ceq $before -and $registryRestored
            if (-not $restored) { throw 'Recent manifest or registry restoration mismatch' }
        }
    } catch {
        $restoreFailure = $_ | Out-String
        $restoreFailure | Set-Content -LiteralPath "$Root\RESTORE_FAILURE.txt" -Encoding UTF8
    } finally {
        if ($script:ExplorerTouched) {
            try { Start-Explorer } catch { $restoreFailure = $_ | Out-String; $restored = $false }
        }
    }
    try {
        $targets = [IO.Path]::GetFullPath("$Root\targets")
        if ($targets -ne "$Root\targets") { throw 'Unsafe target cleanup path' }
        if ($backupReady -and (Test-Path -LiteralPath $targets)) {
            Remove-Item -LiteralPath $targets -Recurse -Force
        }
        $checks = @()
        if (Test-Path -LiteralPath "$Root\checks.tsv") {
            $checks = @(Import-Csv -LiteralPath "$Root\checks.tsv" -Delimiter "`t" |
                ForEach-Object {
                    [pscustomobject]@{ id = $_.id; ok = $_.ok -eq 'true'; detail = $_.detail }
                })
        }
        $completed = @($checks | Where-Object { $_.id -eq 'suite.completed' -and $_.ok }).Count -eq 1
        $unitChecks = 0
        if (Test-Path -LiteralPath "$Root\unit-tests.log") {
            $unitLog = Get-Content -LiteralPath "$Root\unit-tests.log" -Raw
            if ($unitLog -match 'test result: ok\. (\d+) passed; 0 failed;') {
                $unitChecks = [int]$Matches[1]
            }
        }
        Write-Json ([ordered]@{
            run_id = $config.run_id
            platform = $config.platform
            exit_code = $exitCode
            completed = $completed
            checks = $checks.Count
            failures = @($checks | Where-Object { -not $_.ok }).Count
            unit_checks = $unitChecks
            unit_passed = $unitChecks -gt 0
            passed = $completed -and $unitChecks -gt 0 -and $exitCode -eq 0 -and
                $null -eq $failure -and $restored
            check_results = $checks
        }) 'result.json'
        Write-Json ([ordered]@{
            run_id = $config.run_id
            backup_ready = $backupReady
            restored = $restored
            restore_check = 'full Recent tree and registry, before Explorer restart'
            failure = $failure
            restore_failure = $restoreFailure
            targets_removed = -not (Test-Path -LiteralPath "$Root\targets")
            remaining_test_processes = @(
                Get-CimInstance Win32_Process `
                    -Filter "Name='quick-access-manager.exe' OR Name='standalone-unit-tests.exe'" |
                    Where-Object { $_.ExecutablePath -ieq $Exe -or $_.ExecutablePath -ieq $UnitExe }
            ).Count
            timestamp = (Get-Date).ToString('o')
        }) 'final-state.json'
        # Live Explorer can normalize uninitialized CFBs after restart. Record
        # that separately; the host MUST also roll back the whole-VM snapshot.
        (Get-Manifest $Recent) |
            Set-Content -LiteralPath "$Root\recent-after-explorer.txt" -Encoding UTF8
        # Compress-Archive silently excludes hidden files such as desktop.ini.
        # Stage outside Root so the archive cannot include its own open file.
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $archiveStage = "$Root.zip"
        [IO.Compression.ZipFile]::CreateFromDirectory($Root, $archiveStage)
        Move-Item -LiteralPath $archiveStage -Destination "$Root\evidence.zip"
    } catch {
        $_ | Out-String | Set-Content -LiteralPath "$Root\ARTIFACT_FAILURE.txt" -Encoding UTF8
    } finally {
        # This means all cleanup attempts ended, NOT that tests/restoration passed.
        Set-Content -LiteralPath "$Root\finished.txt" -Value $config.run_id
    }
}
