$PackageName = "uninstall_Desktop_Shortcuts"
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

foreach ($shortcut in $shortcuts) {
    $name = $shortcut.Name
    $shortcutPath = Join-Path -Path $commonDesktop -ChildPath "$name.lnk"
    $iconPath = Join-Path -Path $iconDirectory -ChildPath "$name.ico"

    # Remove the shortcut if it exists
    if (Test-Path $shortcutPath) {
        Remove-Item -Path $shortcutPath -Force
        Write-Host "Removed shortcut: $shortcutPath"
    }

    # Remove the icon if it exists
    if (Test-Path $iconPath) {
        Remove-Item -Path $iconPath -Force
        Write-Host "Removed icon: $iconPath"
    }
}

# Delete the icons folder if it exists and is empty
if (Test-Path $iconDirectory) {
    $remainingFiles = Get-ChildItem -Path $iconDirectory -File
    if ($remainingFiles.Count -eq 0) {
        Remove-Item -Path $iconDirectory -Force -Recurse
        Write-Host "Removed empty icons folder: $iconDirectory"
    } else {
        Write-Host "Icons folder not removed as it still contains files."
    }
}

Write-Host "All shortcuts and icons listed in the Excel file have been removed."
