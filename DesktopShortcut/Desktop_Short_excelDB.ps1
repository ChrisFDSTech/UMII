$PackageName = "Desktop_Short_excelDB"
$Version = "1"

# Check if ImportExcel module is installed, if not, install it
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..."
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Import the ImportExcel module
Import-Module ImportExcel

# Download the Excel file from GitHub
$excelUrl = "https://github.com/ChrisFDSTech/UMII/raw/main/DesktopShortcut/DB/shortcuts.xlsx"
$excelPath = Join-Path $env:TEMP "shortcuts.xlsx"
Invoke-WebRequest -Uri $excelUrl -OutFile $excelPath

$commonDesktop = [Environment]::GetFolderPath("CommonDesktopDirectory")
$iconDirectory = Join-Path -Path $env:ProgramData -ChildPath "FDS\Icons"

# Ensure the icon directory exists
if (-not (Test-Path $iconDirectory)) {
    New-Item -ItemType Directory -Path $iconDirectory -Force | Out-Null
}

function Is-WebAppShortcut($shortcutPath) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $targetPath = $shortcut.TargetPath
    return $targetPath -like "*msedge.exe*" -and ($targetPath -like "*--app=*" -or $targetPath -like "*--profile-directory=*")
}

# Import the Excel file
$shortcuts = Import-Excel -Path $excelPath

foreach ($shortcut in $shortcuts) {
    $name = $shortcut.Name
    $url = $shortcut.URL
    $iconUrl = $shortcut.IconPath

    # Replace %computername% with actual computer name
    $name = $name -replace '%computername%', $env:COMPUTERNAME

    $shortcutPath = Join-Path -Path $commonDesktop -ChildPath "$name.lnk"
    $iconPath = Join-Path -Path $iconDirectory -ChildPath "$name.ico"

    # Download icon
    Invoke-WebRequest -Uri $iconUrl -OutFile $iconPath

    # Check if an existing Web App shortcut with the same name exists
    if (Test-Path $shortcutPath) {
        if (Is-WebAppShortcut $shortcutPath) {
            Write-Host "Replacing Web App shortcut for $name with regular URL shortcut"
            Remove-Item $shortcutPath -Force
        }
    }

    # Create or update the shortcut
    $shell = New-Object -ComObject WScript.Shell
    $shortcutFile = $shell.CreateShortcut($shortcutPath)
    $shortcutFile.TargetPath = $url
    $shortcutFile.IconLocation = $iconPath
    $shortcutFile.Save()

    Write-Host "Created/Updated shortcut for $name"
}

Write-Host "All shortcuts have been created/updated."

# Clean up the downloaded Excel file
Remove-Item $excelPath -Force
