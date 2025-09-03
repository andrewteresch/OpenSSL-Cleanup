# Set execution policy to allow all scripts in this session
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process -Force

# Define the target Git version and installer URL
$targetVersion = "2.51.0.windows.1"
$installerUrl = "https://github.com/git-for-windows/git/releases/download/v2.51.0.windows.1/Git-2.51.0-64-bit.exe"
$installerPath = "$env:TEMP\Git-2.51.0-64-bit.exe"

# Check if Git is installed
$gitPath = (Get-Command git -ErrorAction SilentlyContinue).Source
if ($gitPath) {
    Write-Output "Git is installed at $gitPath"
    $currentVersion = (& git --version).Trim()
    Write-Output "Current Git version: $currentVersion"
    if ($currentVersion -match $targetVersion) {
        Write-Output "Git is already at version $targetVersion. No installation needed."
        exit 0
    } elseif ($currentVersion -match "git version (\d+\.\d+\.\d+)") {
        $verNum = $matches[1]
        if ([version]$verNum -lt [version]"2.51.0") {
            Write-Output "Git version is less than $targetVersion. Proceeding with update."
            Write-Output "Downloading Git 2.51.0 installer..."
            Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing
            Write-Output "Download complete."
            Write-Output "Running silent installation of Git 2.51.0..."
            Start-Process -FilePath $installerPath -ArgumentList "/VERYSILENT", "/NORESTART" -Wait
            Write-Output "Installation complete."
            $updatedVersion = (& git --version).Trim()
            Write-Output "Old Git version: $currentVersion"
            Write-Output "Updated Git version: $updatedVersion"
            exit 0
        } else {
            Write-Output "Git version is newer than 2.51. No installation needed."
            exit 0
        }
    } else {
        Write-Output "Could not parse Git version. No installation performed."
        exit 1
    }
} else {
    Write-Output "Git is not installed. No installation will be performed."
    exit 1
}