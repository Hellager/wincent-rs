# !!! This script may be fully stucked on some system, only close window will stop the process

$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$shell = New-Object -ComObject Shell.Application

$folder = $PSScriptRoot

$userInput = Read-Host "Try pin folder '$folder' to frequent folder, confirm to continue('y' to proceed)"
if ($userInput -eq 'y' -or $userInput -eq 'yes') {
    # `pintohome` is a toggle function, which means if you run this twice, the folder will be unpinned
    $shell.NameSpace("$folder").Self.InvokeVerb("pintohome")

    $FrequentFolders = $shell.Namespace('shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}').Items();
    
    $isPinned = $false
    foreach ($item in $FrequentFolders) {
        if ($item.Path -eq $folder) {
            Write-Output "pin $folder success"
            $isPinned = $true
            break
        }
    }

    if ($isPinned) {
        $userInput = Read-Host "Do you want to unpin '$folder' from frequent folder? ('y' to proceed)"
        if ($userInput -eq 'y' -or $userInput -eq 'yes') {
            foreach ($item in $FrequentFolders) {
                if ($item.Path -eq $folder) {
                    $item.InvokeVerb('unpinfromhome')
                    break
                }
            }
        }
    }
} else {
    Write-Output "Stop proceeding, exiting script."
    exit
}
