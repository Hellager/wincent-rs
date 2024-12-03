$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
$shellApplication = New-Object -ComObject Shell.Application;
$windows = $shellApplication.Windows();
$windows | ForEach-Object { 
    Write-Host "Refreshing window: $($_.LocationName)"
    $_.Refresh() 
}