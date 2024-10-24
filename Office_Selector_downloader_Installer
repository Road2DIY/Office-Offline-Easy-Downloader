# Set the working directory to ODT folder
$odtDirectory = "C:\ODT"
$odtSetupUrl = "https://download.microsoft.com/download/F/2/6/F26A4321-1FB2-4C80-9C15-B9CFC1A7C8A8/OfficeDeploymentTool.exe"
$odtSetupPath = "$odtDirectory\OfficeDeploymentTool.exe"

# Create ODT directory if it doesn't exist
if (-not (Test-Path -Path $odtDirectory)) {
    New-Item -Path $odtDirectory -ItemType Directory
}

# Function to download the ODT setup file
function Download-ODTSetup {
    param (
        [string]$url,
        [string]$outputPath
    )

    Write-Host "Downloading Office Deployment Tool (ODT)..."
    Invoke-WebRequest -Uri $url -OutFile $outputPath
    Write-Host "Downloaded ODT setup to $outputPath."
}

# Function to extract ODT setup file
function Extract-ODTSetup {
    param (
        [string]$odtExePath,
        [string]$targetDir
    )

    Write-Host "Extracting ODT setup..."
    Start-Process -FilePath $odtExePath -ArgumentList "/extract:$targetDir" -Wait
    Write-Host "ODT setup extracted to $targetDir."
}

# Check if ODT setup.exe exists, if not download and extract
if (-not (Test-Path "$odtDirectory\setup.exe")) {
    Download-ODTSetup -url $odtSetupUrl -outputPath $odtSetupPath
    Extract-ODTSetup -odtExePath $odtSetupPath -targetDir $odtDirectory
} else {
    Write-Host "ODT setup.exe already exists."
}

# Function to display options and get user's choice
function Get-UserChoice {
    param (
        [string]$message,
        [array]$options
    )
    Write-Host $message
    for ($i = 0; $i -lt $options.Length; $i++) {
        Write-Host "$($i + 1). $($options[$i])"
    }
    $choice = Read-Host "Please enter your choice (1-$($options.Length))"
    if ($choice -ge 1 -and $choice -le $options.Length) {
        return $options[$choice - 1]
    } else {
        Write-Host "Invalid selection. Please try again."
        return Get-UserChoice -message $message -options $options
    }
}

# Function to create configuration XML
function Create-ConfigFile {
    param (
        [string]$officeVersion,
        [string]$arch,
        [string]$channel,
        [string]$outputFile
    )
    
    # Correct the architecture value (32 or 64)
    $archValue = if ($arch -eq "x64") { "64" } else { "32" }
    
    $xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="$archValue" Channel="$channel">
    <Product ID="$officeVersion">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="C:\ODT" />
</Configuration>
"@
    
    # Save XML content to the config file
    $xmlContent | Out-File -FilePath $outputFile -Encoding UTF8
    Write-Host "Configuration file $outputFile created."
}

# Office versions
$officeVersions = @("Office 2019", "Office 2021", "Office 2024", "Office 365")

# Architecture options
$architectures = @("x64", "x32")

# Map version names to Office IDs and update channels
$officeIDMap = @{
    "Office 2019" = "ProPlus2019Volume"
    "Office 2021" = "ProPlus2021Volume"
    "Office 2024" = "ProPlus2024Volume"
    "Office 365"  = "O365ProPlusRetail"
}

$channelMap = @{
    "Office 2019" = "PerpetualVL2019"
    "Office 2021" = "PerpetualVL2021"
    "Office 2024" = "PerpetualVL2021"
    "Office 365"  = "Current"
}

# Get user's office version choice
$selectedVersion = Get-UserChoice -message "Select the Office version you want to install:" -options $officeVersions

# Get user's architecture choice
$selectedArch = Get-UserChoice -message "Select the architecture (x64 or x32):" -options $architectures

# Generate the Office Product ID
$officeProductID = $officeIDMap[$selectedVersion]

# Get the correct update channel
$updateChannel = $channelMap[$selectedVersion]

# Generate the configuration file name
$configFileName = "config-$($selectedVersion.Replace(' ',''))-$selectedArch.xml"
$configFilePath = "$odtDirectory\$configFileName"

# Create the configuration file
Create-ConfigFile -officeVersion $officeProductID -arch $selectedArch -channel $updateChannel -outputFile $configFilePath

# Check if the configuration file exists
if (Test-Path $configFilePath) {
    Write-Host "Downloading and installing $selectedVersion ($selectedArch)..."
    # Run the Office Deployment Tool with the appropriate config file
    Start-Process -FilePath "$odtDirectory\setup.exe" -ArgumentList "/configure $configFilePath" -Wait
    Write-Host "$selectedVersion installation completed."
} else {
    Write-Host "Failed to create the configuration file."
}

# End of script
