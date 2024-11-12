$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$shell = New-Object -ComObject Shell.Application
$files = $shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items() | Where-Object { $_.IsFolder -eq $false }

# Get the first file
$firstFile = $files | Select-Object -First 1

# Print the Name and Path
if ($firstFile) {
	$userInput = Read-Host "Try remove file '$($firstFile.Name)' from recent, confirm to continue('y' to proceed)"
	if ($userInput -eq 'y' -or $userInput -eq 'yes') {
		$firstFile.InvokeVerb("remove");
	} else {
		Write-Output "Stop proceeding, exiting script."
		exit
	}
} else {
    Write-Output "No files found."
}