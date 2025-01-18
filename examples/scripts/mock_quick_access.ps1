# Function to generate random text
function Get-RandomText {
    $words = @("test", "quick", "access", "windows", "file", "folder", "content", "random", "text", "data")
    $length = Get-Random -Minimum 5 -Maximum 15
    $text = ""
    for ($i = 0; $i -lt $length; $i++) {
        $text += $words[(Get-Random -Maximum $words.Length)] + " "
    }
    return $text.Trim()
}

# Add type for SHAddToRecentDocs
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class RecentDocs {
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern void SHAddToRecentDocs(uint flags, string path);
    }
"@

# Create test directory
$testRoot = Join-Path $PSScriptRoot "wincent-test"
if (!(Test-Path $testRoot)) {
    New-Item -ItemType Directory -Path $testRoot | Out-Null
}

# Create 25 text files with random content
Write-Host "Creating test files..."
1..25 | ForEach-Object {
    $fileName = "test_file_$_.txt"
    $filePath = Join-Path $testRoot $fileName
    $content = Get-RandomText
    Set-Content -Path $filePath -Value $content
    
    # Open and close with Notepad
    Start-Process "notepad.exe" -ArgumentList $filePath
    Write-Host "Created and opened file: $fileName"
}

Get-Process notepad | Where-Object { $_.MainWindowTitle -like "*Notepad*" } | ForEach-Object { Stop-Process -Id $_.Id }

# Create and pin 10 folders
Write-Host "`nCreating and pinning test folders..."
1..10 | ForEach-Object {
    $folderName = "test_folder_$_"
    $folderPath = Join-Path $testRoot $folderName
    
    # Create folder
    if (!(Test-Path $folderPath)) {
        New-Item -ItemType Directory -Path $folderPath | Out-Null
    }
    
    # Pin folder to Quick Access
    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.Namespace($folderPath)
    $folder.Self.InvokeVerb("pintohome")
    
    Write-Host "Created and pinned folder: $folderName"
}

Write-Host "`nTest environment setup complete!"
Write-Host "Test files location: $testRoot"

$shellApplication = New-Object -ComObject Shell.Application;
$windows = $shellApplication.Windows();
$windows | ForEach-Object { 
    $_.Refresh() 
}

Write-Host "Refresh windows complete"
