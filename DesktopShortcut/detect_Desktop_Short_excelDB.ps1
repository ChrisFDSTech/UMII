$PackageName = "detect_Desktop_Short_excelDB"
$Version = "1"

# Check if ImportExcel module is installed, if not, install it
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..."
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Import the ImportExcel module
Import-Module ImportExcel

$excelPath = "C:\Path\To\Your\shortcuts.xlsx"
$commonDesktop = [Environment]::GetFolderPath("CommonDesktopDirectory")
$iconDirectory = Join-Path -Path $env:ProgramData -ChildPath "FDS\Icons"

# Import the Excel file
$shortcuts = Import-Excel -Path $excelPath

$allShortcutsCorrect = $true

foreach ($shortcut in $shortcuts) {
    $name = $shortcut.Name
    $url = $shortcut.URL
    $iconPath = Join-Path -Path $iconDirectory -ChildPath "$name.ico"
    $shortcutPath = Join-Path -Path $commonDesktop -ChildPath "$name.lnk"

    if (Test-Path $shortcutPath) {
        $shell = New-Object -ComObject WScript.Shell
        $existingShortcut = $shell.CreateShortcut($shortcutPath)

        if ($existingShortcut.TargetPath -ne $url -or $existingShortcut.IconLocation -ne $iconPath) {
            Write-Host "Shortcut $name exists but has incorrect properties."
            $allShortcutsCorrect = $false
        }
    } else {
        Write-Host "Shortcut $name does not exist."
        $allShortcutsCorrect = $false
    }
}

if ($allShortcutsCorrect) {
    Write-Host "All shortcuts exist and have correct properties."
    exit 0
} else {
    Write-Host "One or more shortcuts are missing or have incorrect properties."
    exit 1
}
