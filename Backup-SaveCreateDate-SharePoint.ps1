# Parameters
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$LocalBackupRoot,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$SubFolderServerRelativeUrl,  # Optional parameter for subfolder

    [string]$LogFilePath,

    [int]$PageSize = 1000  # Optional page size for batch processing
)

# Set default log file path if not provided
if (-not $LogFilePath) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogDirectory = Join-Path $LocalBackupRoot "Logs"
    if (!(Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory | Out-Null
    }
    $LogFilePath = Join-Path $LogDirectory "BackupLog_$timestamp.txt"
}

# Logging function with an option to skip console output
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$NoConsole  # Suppress console output if this switch is used
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    if (-not $NoConsole) {
        Write-Host $logEntry
    }
    $logEntry | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
}

# Start logging
Write-Log "Backup script started."

# Install PnP.PowerShell Module if not already installed
if (!(Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Log "Installing PnP.PowerShell module."
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}

# Connect to SharePoint Online Site using the registered Entra ID application
try {
    Write-Log "Connecting to SharePoint Online site: $SiteUrl"
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive
    Write-Log "Successfully connected to SharePoint Online."
}
catch {
    Write-Log "Error connecting to SharePoint Online: $_" "ERROR"
    exit 1
}

# Get the site object to use in URL conversions
try {
    $Site = Get-PnPWeb
    Write-Log "Retrieved Site Object: $($Site.Title)"
}
catch {
    Write-Log "Error retrieving Site Object: $_" "ERROR"
    Disconnect-PnPOnline
    exit 1
}

# Function to download files and folders recursively with pagination
function Download-Files {
    param (
        [string]$ServerRelativeUrl,
        [string]$LocalPath
    )
    Write-Log "Processing: $ServerRelativeUrl"

    # Try to get the folder
    try {
        Write-Log "Attempting to retrieve folder: $ServerRelativeUrl"
        $folder = Get-PnPFolder -ServerRelativeUrl $ServerRelativeUrl -ErrorAction Stop
        $isFolder = $true
        Write-Log "Folder found: $($folder.Name)"
    }
    catch {
        Write-Log "Could not retrieve folder: $ServerRelativeUrl - Treating as document library" "WARNING"
        $isFolder = $false
    }

    if ($isFolder) {
        # Process SharePoint Folder as before (omitted for brevity)
    }
    else {
        # Process Document Library in Batches
        Write-Log "Processing Document Library at '$ServerRelativeUrl' in batches of $PageSize"

        # Get the list (Document Library)
        try {
            $list = Get-PnPList | Where-Object { $_.RootFolder.ServerRelativeUrl -eq $ServerRelativeUrl } | Select-Object -First 1
            if ($null -eq $list) {
                throw "No list found at '$ServerRelativeUrl'"
            }
            Write-Log "Found Document Library: $($list.Title)" -NoConsole
        }
        catch {
            Write-Log "Error retrieving Document Library at '$ServerRelativeUrl': $_" "ERROR"
            return
        }

        # Get all items in the library in batches
        $listItems = $null
        $currentBatch = 1
        do {
            try {
                Write-Log "Retrieving batch ${currentBatch} from Document Library '$($list.Title)'" -NoConsole
                $listItems = Get-PnPListItem -List $list -PageSize $PageSize -Fields FileLeafRef,FileRef,FSObjType,Created,Modified
                Write-Log "Batch ${currentBatch}: Number of items retrieved: $($listItems.Count)" -NoConsole

                foreach ($item in $listItems) {
                    $fileRef = $item["FileRef"]
                    $fileLeafRef = $item["FileLeafRef"]
                    $isFolder = $item["FSObjType"] -eq 1
                    $createdDate = $item["Created"]
                    $modifiedDate = $item["Modified"]

                    $relativePath = if ($Site.ServerRelativeUrl -ne "/") {
                        $fileRef.Substring($Site.ServerRelativeUrl.Length).TrimStart('/')
                    }
                    else {
                        $fileRef.TrimStart('/')
                    }

                    $localItemPath = Join-Path $LocalPath $relativePath

                    if ($isFolder) {
                        Write-Log "Processing folder: $fileRef" -NoConsole

                        if (!(Test-Path $localItemPath)) {
                            New-Item -ItemType Directory -Path $localItemPath | Out-Null
                        }

                        # Set folder creation and last modified date
                        try {
                            (Get-Item $localItemPath).CreationTime = $createdDate
                            (Get-Item $localItemPath).LastWriteTime = $modifiedDate
                            Write-Log "Set CreatedDate and LastModifiedDate for folder '$localItemPath' to '$createdDate' and '$modifiedDate'" -NoConsole
                        }
                        catch {
                            Write-Log "Error setting CreatedDate and LastModifiedDate for folder '$localItemPath': $_" "ERROR"
                        }
                    }
                    else {
                        # Log file operations only to the log file, not to the console
                        Write-Log "Processing file: $fileRef" -NoConsole
                        $localFolderPath = Split-Path $localItemPath -Parent

                        if (!(Test-Path $localFolderPath)) {
                            New-Item -ItemType Directory -Path $localFolderPath -Force | Out-Null
                        }

                        try {
                            # Download the file
                            Get-PnPFile -Url $fileRef -Path $localFolderPath -FileName $fileLeafRef -AsFile -Force
                            Write-Log "Downloaded file: $fileRef" -NoConsole

                            # Set file creation and last modified date
                            try {
                                $localFilePath = Join-Path $localFolderPath $fileLeafRef
                                (Get-Item $localFilePath).CreationTime = $createdDate
                                (Get-Item $localFilePath).LastWriteTime = $modifiedDate
                                Write-Log "Set CreatedDate and LastModifiedDate for file '$localFilePath' to '$createdDate' and '$modifiedDate'" -NoConsole
                            }
                            catch {
                                Write-Log "Error setting CreatedDate and LastModifiedDate for file '$localFilePath': $_" "ERROR"
                            }
                        }
                        catch {
                            Write-Log "Error downloading file '$fileRef': $_" "ERROR"
                        }
                    }
                }
                $currentBatch++
            }
            catch {
                Write-Log "Error retrieving batch ${currentBatch} from Document Library '$($list.Title)': $_" "ERROR"
                break
            }

        } while ($listItems -ne $null -and $listItems.Count -eq $PageSize)  # Continue if more items exist
    }
}

# Main Execution
if ($SubFolderServerRelativeUrl) {
    # Start downloading from the specified subfolder or library
    Write-Log "Backing up from: $SubFolderServerRelativeUrl"

    # Ensure the path starts with '/'
    if (-not $SubFolderServerRelativeUrl.StartsWith('/')) {
        $SubFolderServerRelativeUrl = '/' + $SubFolderServerRelativeUrl
    }

    # Start downloading
    Download-Files -ServerRelativeUrl $SubFolderServerRelativeUrl -LocalPath $LocalBackupRoot
}
else {
    Write-Log "No subfolder specified, exiting script." "ERROR"
}

# Disconnect from SharePoint Online
Disconnect-PnPOnline
Write-Log "Backup script completed."
