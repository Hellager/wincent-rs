# Windows PowerShell 5.1. Invoked only by the guarded VM guest runner.
param([Parameter(Mandatory = $true)][string]$Root)

$ErrorActionPreference = 'Stop'
$UnitExe = Join-Path $Root 'library-unit-tests.exe'
$SuiteExe = Join-Path $Root 'quick-access-manager.exe'
$UnitLog = Join-Path $Root 'unit-tests.log'
$UnitList = Join-Path $Root 'unit-tests-list.txt'
$UnitCount = Join-Path $Root 'unit-tests-count.txt'
$SuiteLog = Join-Path $Root 'test.log'

try {
    $listed = @(& $UnitExe '--list' 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "Listing library tests failed: $LASTEXITCODE" }
    $tests = @(
        $listed | ForEach-Object {
            if ([string]$_ -match '^(?<name>.+): test$') { $Matches.name }
        } | Where-Object { $_ -and $_ -notmatch '^destlist::' }
    )
    if ($tests.Count -eq 0) { throw 'No non-destlist library tests were discovered' }
    $tests | Set-Content -LiteralPath $UnitList -Encoding UTF8
    Set-Content -LiteralPath $UnitLog -Value '' -Encoding UTF8

    $passed = 0
    foreach ($test in $tests) {
        "=== $test ===" | Add-Content -LiteralPath $UnitLog -Encoding UTF8
        $command = "`"$UnitExe`" `"$test`" --exact --test-threads=1 --nocapture >> `"$UnitLog`" 2>&1"
        & cmd.exe /d /c $command
        if ($LASTEXITCODE -ne 0) { throw "Library test failed: $test" }
        $passed++
    }
    Set-Content -LiteralPath $UnitCount -Value $passed -Encoding ASCII

    $command = "`"$SuiteExe`" --ignored --exact quick_access_manager --nocapture >> `"$SuiteLog`" 2>&1"
    & cmd.exe /d /c $command
    exit $LASTEXITCODE
} catch {
    $_ | Out-String | Add-Content -LiteralPath $UnitLog -Encoding UTF8
    exit 1
}
