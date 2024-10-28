# SharePoint Site URL
$siteUrl = 
$ClientId = 
$TenantId = 

# Document Library Name
$basePath = 
$txtFilePath = 

# Read the text file into a list (array)

$folderNames = Get-Content -Path $txtFilePath
$sourceFolderRelativeUrls = @()
# Remove empty lines and comments (lines starting with #)
foreach ($folderName in $folderNames) {
    # Trim any leading/trailing whitespace
    $folderName = $folderName.Trim()

    # Skip empty lines or comments (lines starting with #)
    if ([string]::IsNullOrEmpty($folderName) -or $folderName.StartsWith("#")) {
        continue
    }

    # Concatenate the base path
    $fullPath = "$basePath$folderName"

    # Add to the list
    $sourceFolderRelativeUrls += $fullPath
}


# Local Download Base Path
$downloadBasePath = 

# Log file base path
$logFileBasePath =

# Import PnP PowerShell Module
Import-Module PnP.PowerShell -ErrorAction Stop

# Ensure the log file directory exists
$logDir = Split-Path -Parent $logFileBasePath
if (!(Test-Path -Path $logDir)) {
    try {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    catch {
        Write-Host "❌ Failed to create log directory: $logDir. Error: $_" -ForegroundColor Red
        exit 1
    }
}

# Initialize log file
Clear-Content -Path $logFileBasePath -ErrorAction SilentlyContinue

# Define Write-Log function
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$NoConsole,
        [ConsoleColor]$Color = [ConsoleColor]::White
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    # Write to log file
    Add-Content -Path $logFileBasePath -Value $logEntry
    # Write to console if $NoConsole is not specified
    if (-not $NoConsole) {
        Write-Host "[$timestamp] $Message" -ForegroundColor $Color
    }
}

# Function to ensure a folder exists locally and set its dates
function Ensure-LocalFolder {
    param (
        [string]$FolderSiteRelativeUrl,
        [string]$LocalFolderPath
    )

    # Retrieve folder properties from SharePoint
    try {
        $folder = Get-PnPFolder -Url $FolderSiteRelativeUrl -Includes TimeCreated, TimeLastModified
    }
    catch {
        Write-Log "❌ Failed to get properties for folder: $FolderSiteRelativeUrl. Error: $_" -Level "ERROR" -Color Red -NoConsole
        return
    }

    # Check if the folder exists locally
    if (Test-Path -Path $LocalFolderPath) {
        # Folder exists, check creation date
        $localFolder = Get-Item $LocalFolderPath
        if ($localFolder.CreationTime -eq $folder.TimeCreated) {
            Write-Log "📁 Folder already exists with the same creation date: $LocalFolderPath" -Color Yellow -NoConsole
            return
        }
        else {
            Write-Log "📁 Folder exists but with different creation date: $LocalFolderPath" -Color Cyan -NoConsole
            # Update the folder's creation and modification dates
            $localFolder.CreationTime = $folder.TimeCreated
            $localFolder.LastWriteTime = $folder.TimeLastModified
            Write-Log "🕒 Updated dates for folder: $LocalFolderPath" -Color Cyan -NoConsole
            return
        }
    }
    else {
        # Folder does not exist, create it
        New-Item -ItemType Directory -Path $LocalFolderPath -Force | Out-Null
        Write-Log "📁 Created folder: $LocalFolderPath" -Color Cyan -NoConsole

        # Set the local folder's creation and modification dates
        (Get-Item $LocalFolderPath).CreationTime = $folder.TimeCreated
        (Get-Item $LocalFolderPath).LastWriteTime = $folder.TimeLastModified
        Write-Log "🕒 Set dates for folder: $LocalFolderPath" -Color Cyan -NoConsole
    }
}

# Function to process a folder and its contents recursively
function ProcessFolder {
    param(
        [string]$sourceFolderSiteRelativeUrl,
        [string]$destinationFolderPath
    )

    Write-Log "Processing folder: $sourceFolderSiteRelativeUrl" -Color Cyan

    # Ensure the local folder exists and set its dates
    Ensure-LocalFolder -FolderSiteRelativeUrl $sourceFolderSiteRelativeUrl -LocalFolderPath $destinationFolderPath

    # Get items in the folder
    try {
        $items = Get-PnPFolderItem -FolderSiteRelativeUrl $sourceFolderSiteRelativeUrl
    }
    catch {
        Write-Log "❌ Failed to get items for folder: $sourceFolderSiteRelativeUrl. Error: $_" -Level "ERROR" -Color Red -NoConsole
        return
    }

    foreach ($item in $items) {
        if ($item.PSObject.TypeNames -contains "Microsoft.SharePoint.Client.Folder") {
            # It's a folder
            $folderName = $item.Name
            $folderServerRelativeUrl = $item.ServerRelativeUrl
            $folderSiteRelativeUrl = $folderServerRelativeUrl.Substring($web.ServerRelativeUrl.Length)

            $localSubFolderPath = Join-Path $destinationFolderPath $folderName

            # Ensure the folder exists and set its dates
            Ensure-LocalFolder -FolderSiteRelativeUrl $folderSiteRelativeUrl -LocalFolderPath $localSubFolderPath

            # Recursively process the subfolder
            ProcessFolder -sourceFolderSiteRelativeUrl $folderSiteRelativeUrl -destinationFolderPath $localSubFolderPath
        }
        elseif ($item.PSObject.TypeNames -contains "Microsoft.SharePoint.Client.File") {
            # It's a file
            $fileName = $item.Name
            $fileServerRelativeUrl = $item.ServerRelativeUrl
            $fileSiteRelativeUrl = $fileServerRelativeUrl.Substring($web.ServerRelativeUrl.Length)

            $localFilePath = Join-Path $destinationFolderPath $fileName

            # Get file properties
            try {
                $file = Get-PnPFile -Url $fileSiteRelativeUrl -AsListItem
                $createdDate = $file.FieldValues["Created"]
                $modifiedDate = $file.FieldValues["Modified"]
            }
            catch {
                Write-Log "❌ Failed to get properties for file: $fileSiteRelativeUrl. Error: $_" -Level "ERROR" -Color Red -NoConsole
                continue
            }

            $downloadFile = $true

            if (Test-Path -Path $localFilePath) {
                # File exists, check creation date
                $localFile = Get-Item $localFilePath
                if ($localFile.CreationTime -eq $createdDate) {
                    Write-Log "📄 File already exists with the same creation date: $localFilePath" -Color Yellow -NoConsole
                    $downloadFile = $false
                }
                else {
                    Write-Log "📄 File exists but with different creation date: $localFilePath" -Color Cyan -NoConsole
                    # Optionally, update the file's creation date or decide to download it again
                    $downloadFile = $false
                }
            }

            if ($downloadFile) {
                try {
                    Get-PnPFile -Url $fileSiteRelativeUrl -Path $destinationFolderPath -FileName $fileName -AsFile -Force
                    Write-Log "✅ Downloaded file: $fileName to $destinationFolderPath" -Color Green -NoConsole

                    # Set the file's creation and modification dates
                    (Get-Item $localFilePath).CreationTime = $createdDate
                    (Get-Item $localFilePath).LastWriteTime = $modifiedDate
                }
                catch {
                    Write-Log "❌ Failed to download file: $fileName. Error: $_" -Level "ERROR" -Color Red -NoConsole
                }
            }
        }
    }

    Write-Log "✅ Completed processing for folder: $sourceFolderSiteRelativeUrl" -Color Green
}

# Connect to SharePoint Online
Write-Log "Connecting to SharePoint site: $siteUrl" -Color Cyan

try {
    Write-Log "Connecting to SharePoint Online..."
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive
    Write-Log "Successfully connected to SharePoint Online."

    # Get the web object to determine the site's server-relative URL
    $web = Get-PnPWeb
}
catch {
    Write-Log "Failed to connect to SharePoint Online: $_" -Level "ERROR" -Color Red
    exit 1
}

# Process each folder in the list
foreach ($sourceFolderRelativeUrl in $sourceFolderRelativeUrls) {
    # Convert server-relative URL to site-relative URL
    $sourceFolderSiteRelativeUrl = $sourceFolderRelativeUrl.Substring($web.ServerRelativeUrl.Length)
    if ($sourceFolderSiteRelativeUrl.StartsWith("/")) {
        $sourceFolderSiteRelativeUrl = $sourceFolderSiteRelativeUrl.Substring(1)
    }

    # Determine the local destination path for each folder
    $folderName = Split-Path $sourceFolderRelativeUrl -Leaf
    $destinationFolderPath = Join-Path $downloadBasePath $folderName

    # Start processing the folder
    ProcessFolder -sourceFolderSiteRelativeUrl $sourceFolderSiteRelativeUrl -destinationFolderPath $destinationFolderPath
}

# Disconnect from SharePoint Online
Disconnect-PnPOnline
Write-Log "✅ All folders and files copied successfully." -Color Green
