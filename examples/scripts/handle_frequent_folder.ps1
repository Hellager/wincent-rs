# !!! This script may be fully stucked on some system, only close window will stop the process

$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$shell = New-Object -ComObject Shell.Application

$folder = $PSScriptRoot

$userInput = Read-Host "Try pin folder '$folder' to frequent folder, confirm to continue('y' to proceed)"
if ($userInput -eq 'y' -or $userInput -eq 'yes') {
    # `pintohome` is a toggle function, which means if you run this twice, the folder will be unpinned
    $shell.NameSpace("$folder").Self.InvokeVerb('pintohome')
} else {
    Write-Output "Stop proceeding, exiting script."
    exit
}
