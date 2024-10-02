# Set the directory for the Office Deployment Tool (ODT)
$ODTDir = "$PSScriptRoot\\ODT"

# Check if the Office Deployment Tool directory exists
if (!(Test-Path -Path $ODTDir)) {
    Write-Host "ODT directory not found. Please ensure the ODT files are located in: $ODTDir"
    exit
}

# Verify that setup.exe exists in the ODT directory
$SetupPath = "$ODTDir\\setup.exe"
if (!(Test-Path -Path $SetupPath)) {
    Write-Host "setup.exe not found in the ODT directory. Please ensure it is present in: $ODTDir"
    exit
}

# Function to create configuration files for Office versions
function Create-ConfigurationFile {
    param (
        [string]$OfficeDir,
        [string]$OfficeClientEdition,
        [string]$OfficeVersion,
        [string]$ProductID
    )

    Write-Host "Creating configuration file for $OfficeVersion ($OfficeClientEdition-bit)..."
    $ConfigContent = @"
<Configuration>
  <Add SourcePath="$OfficeDir" OfficeClientEdition="$OfficeClientEdition" Channel="$OfficeVersion">
    <Product ID="$ProductID">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="True" />
</Configuration>
"@
    $ConfigContent | Out-File -FilePath "$OfficeDir\\configuration_$OfficeClientEdition.xml" -Encoding UTF8
}

# Function to create installation scripts for Office versions
function Create-InstallScript {
    param (
        [string]$OfficeDir,
        [string]$OfficeClientEdition,
        [string]$OfficeVersion
    )

    Write-Host "Creating installation script for $OfficeVersion ($OfficeClientEdition-bit)..."
    $InstallScriptContent = @"
@echo off
echo Installing $OfficeVersion ($OfficeClientEdition-bit)...
cd `"$ODTDir`"
setup.exe /configure `"$OfficeDir\\configuration_$OfficeClientEdition.xml`"
pause
"@
    $InstallScriptContent | Out-File -FilePath "$OfficeDir\\Install_$OfficeVersion_$OfficeClientEdition.bat" -Encoding ASCII
}

# Office versions and products to be processed
$OfficeVersions = @(
    @{Version = "PerpetualVL2019"; ProductID = "ProPlus2019Volume"},
    @{Version = "PerpetualVL2021"; ProductID = "ProPlus2021Volume"},
    @{Version = "PerpetualVL2024"; ProductID = "ProPlus2024Volume"}
)

# Loop through each Office version and create download and installation scripts for 32-bit and 64-bit
foreach ($Office in $OfficeVersions) {
    foreach ($OfficeClientEdition in @("32", "64")) {
        # Set directories for each version and architecture
        $OfficeDir = "$PSScriptRoot\\Office$($Office.Version)_$OfficeClientEdition" + "bit"
        if (!(Test-Path -Path $OfficeDir)) {
            New-Item -Path $OfficeDir -ItemType Directory
        }

        # Create configuration files
        Create-ConfigurationFile -OfficeDir $OfficeDir -OfficeClientEdition $OfficeClientEdition -OfficeVersion $Office.Version -ProductID $Office.ProductID

        # Download Office
        Write-Host "Downloading $($Office.Version) ($OfficeClientEdition-bit)..."
        try {
            Start-Process -FilePath "$SetupPath" -ArgumentList "/download `"$OfficeDir\\configuration_$OfficeClientEdition.xml`"" -Wait -ErrorAction Stop
        } catch {
            Write-Host "Failed to download $($Office.Version) ($OfficeClientEdition-bit)."
            exit
        }

        # Create installation scripts
        Create-InstallScript -OfficeDir $OfficeDir -OfficeClientEdition $OfficeClientEdition -OfficeVersion $Office.Version
    }
}

Write-Host "All tasks completed. The download and installation scripts have been created."
