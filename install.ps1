#run clean up job remove all version.
$ErrorActionPreference = "Silentlycontinue"

#Start logging
#Not it will remove any version of 7zip before upgrading the new one.
Start-Transcript -path "c:\temp\zip_updater.log"

write-output "Running Cleanup Job"
# Check the registry for installed 7-Zip versions
$RegistryPaths = @(
    "HKLM:\SOFTWARE\7-Zip",
    "HKLM:\SOFTWARE\Wow6432Node\7-Zip"
)

$foundVersions = @()

foreach ($path in $RegistryPaths) {
    if (Test-Path $path) {
        $version = (Get-ItemProperty -Path $path).Version
        Write-Output "7-Zip found in registry: Version $version"
        $foundVersions += $version
    }
}

# Search in 'Program Files' and 'Program Files (x86)' directories
$InstallPaths = @(
    "${env:ProgramFiles}\7-Zip\",
    "${env:ProgramFiles(x86)}\7-Zip\"
)

foreach ($dir in $InstallPaths) {
    if (Test-Path $dir) {
        Write-Output "7-Zip found at: $dir"
        $foundVersions += $dir
    }
}

# Uninstall found versions
foreach ($version in $foundVersions) {
    Write-Output "Attempting to remove 7-Zip version at: $version"
    
    # Locate uninstaller using Windows Installer
    $app = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE '7-Zip%'"
    if ($app) {
        $app.Uninstall()
        Write-Output "Uninstalled: $app.Name"
    } else {
        Write-Output "No uninstaller found for version: $version"
    }
}

Write-output "Cleanup Job completed!"


write-output "Update Started"
$msipath = "$PSScriptRoot\7zipinstall.msi"
$installProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /qn /norestart" -NoNewWindow -Wait -PassThru

# Capture the exit code
$exitCode = $installProcess.ExitCode

# Determine installation status
if ($exitCode -eq 0) {
    Write-Output "Installation successfully completed."
} elseif ($exitCode -eq $null) {
    Write-Output "Could not capture the install code."
} else {
    Write-Output "Installation failed with exit code: $exitCode"
}

Write-Output "Installation process finish please look at the logs for review"
 


#stop log
Stop-Transcript
