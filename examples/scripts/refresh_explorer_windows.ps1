$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
$shellApplication = New-Object -ComObject Shell.Application;
$windows = $shellApplication.Windows();
$count = $windows.Count();

foreach( $i in 0..($count-1) ) {
	$item = $windows.Item( $i )
	$item.Refresh() 
}