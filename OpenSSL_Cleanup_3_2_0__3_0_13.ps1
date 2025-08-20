# Set execution policy to allow script execution
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# OpenSSL DLL Cleanup Script for Office16 and Salesforce ODBC paths
$dryRun = $false  # Set to $true for testing without deletion

# Define DLL paths and their corresponding target versions
$dllTargets = @(
    @{ Path = "C:\Program Files\Microsoft Office\root\Office16\libcrypto-3-x64.dll"; Version = "3.2.0" },
    @{ Path = "C:\Program Files\Microsoft Office\root\Office16\ODBC Drivers\Salesforce\libcrypto-3-x64.dll"; Version = "3.2.0" },
    @{ Path = "C:\Program Files\Microsoft Office\root\Office16\ODBC Drivers\Salesforce\lib\openssl64.dlla\libssl-3-x64.dll"; Version = "3.0.13" },
    @{ Path = "C:\Program Files\Microsoft Office\root\Office16\ODBC Drivers\Salesforce\lib\openssl64.dlla\libcrypto-3-x64.dll"; Version = "3.0.13" },
    @{ Path = "C:\Program Files\Microsoft Office\root\Office16\ODBC Drivers\Salesforce\lib\libcurl64.dlla\openssl64.dlla\libssl-3-x64.dll"; Version = "3.0.13" },
    @{ Path = "C:\Program Files\Microsoft Office\root\Office16\ODBC Drivers\Salesforce\lib\libcurl64.dlla\openssl64.dlla\libcrypto-3-x64.dll"; Version = "3.0.13" }
)

Write-Host "`nStarting OpenSSL DLL Cleanup Script (Dry Run Mode: $dryRun)"

$removedFiles = 0
$uninstalledEntries = 0

foreach ($dll in $dllTargets) {
    $dllPath = $dll.Path
    $targetVersion = $dll.Version
    if (Test-Path $dllPath) {
        $versionInfo = (Get-Item $dllPath).VersionInfo.ProductVersion
        # Normalize version strings for comparison
        $normalizedTargetVersion = $targetVersion.TrimEnd(".0")
        $normalizedFileVersion = $versionInfo.TrimEnd(".0")
        Write-Host "Found: $dllPath (Version: $versionInfo)"
        
        if ($normalizedFileVersion -eq $normalizedTargetVersion) {
            Write-Host "Target version matched: $targetVersion"
            if ($dryRun) {
                Write-Host "Dry run: Would remove $dllPath"
            } else {
                try {
                    Remove-Item -Path $dllPath -Force
                    Write-Host "Removed: $dllPath"
                    $removedFiles++
                } catch {
                    Write-Host "Failed to remove $dllPath - $_"
                }
            }
        } else {
            Write-Host "Skipping: Version mismatch ($versionInfo)"
        }
    } else {
        Write-Host "File not found: $dllPath"
    }
}

# Check registry for related OpenSSL and Salesforce entries for both versions
Write-Host "`nChecking registry for OpenSSL and Salesforce entries..."
$registryPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$targetVersions = @("3.2.0", "3.0.13")

foreach ($regPath in $registryPaths) {
    $entries = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue |
        Where-Object { 
            ($_.DisplayName -like "*OpenSSL*" -or $_.DisplayName -like "*Salesforce*") -and 
            ($targetVersions -contains ($_.DisplayVersion.TrimEnd(".0")))
        }

    foreach ($entry in $entries) {
        Write-Host "Registry entry found: $($entry.DisplayName) version $($entry.DisplayVersion)"
        if ($dryRun) {
            Write-Host "Dry run: Would uninstall $($entry.DisplayName)"
        } else {
            try {
                Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $($entry.PSChildName) /quiet" -Wait
                Write-Host "Uninstalled: $($entry.DisplayName)"
                $uninstalledEntries++
            } catch {
                Write-Host "Failed to uninstall $($entry.DisplayName) - $_"
            }
        }
    }
}

# Clean the Office 365 update cache
Write-Host "`nCleaning Office 365 update cache..."
$updateCachePath = "C:\Program Files\Microsoft Office\updates\download\packagefiles"
if (Test-Path $updateCachePath) {
    if ($dryRun) {
        Write-Host "Dry run: Would clear update cache at $updateCachePath"
    } else {
        try {
            Remove-Item -Path "$updateCachePath\*" -Recurse -Force
            Write-Host "Office update cache cleared."
        } catch {
            Write-Host "Failed to clear update cache - $_"
        }
    }
} else {
    Write-Host "Update cache path not found."
}

Write-Host "`nSummary:"
Write-Host "Files removed: $removedFiles"
Write-Host "Registry entries uninstalled: $uninstalledEntries"
Write-Host "Script completed."