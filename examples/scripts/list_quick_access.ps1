$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
$shell = New-Object -ComObject Shell.Application;
$QuickAccess = $shell.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}').Items();

$nonExistentPaths = @()
$index = 1
Write-Output ''
Write-Output '--------------------------- Full Quick Access ---------------------------' 
foreach ($item in $QuickAccess) {
    Write-Output "$index. $($item.Path)"
	$index++

    if (-not (Test-Path $item.Path)) {
        $nonExistentPaths += $item.Path
    }
}

$FrequentFolders = $shell.Namespace('shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}').Items();
$index = 1
Write-Output ''
Write-Output '--------------------------- Frequent Folders ---------------------------' 
foreach ($item in $FrequentFolders) {
    Write-Output "$index. $($item.Path)"
	$index++
}

$RecentFiles = $shell.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}').Items() | Where-Object {$_.IsFolder -eq $false};
$index = 1
Write-Output ''
Write-Output '--------------------------- Recent Files ---------------------------' 
foreach ($item in $RecentFiles) {
		Write-Output "$index. $($item.Path)"
		$index++
}

if ($nonExistentPaths.Count -gt 0) {
    Write-Output ''
	Write-Output '--------------------------- Not Exist Path ---------------------------' 
    foreach ($item in $nonExistentPaths) {
		Write-Output "$index. $($item.Path)"
		$index++
	}
}
