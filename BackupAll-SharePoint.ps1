# Requires PowerShell 7 or later

[CmdletBinding()]
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

    [string]$LogFilePath
)

# Ensure PowerShell 7 or later is being used
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "This script requires PowerShell version 7 or later."
    exit 1
}

# Import necessary .NET namespaces for concurrent queues
Add-Type -AssemblyName System.Collections.Concurrent

# Initialize concurrent queues for logs and errors
$logQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$errorQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# Function to enqueue logs with optional console output
function Enqueue-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$Output
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    $logQueue.Enqueue($logEntry)
    if ($Output) {
        if ($Level -eq "ERROR") {
            Write-Host $logEntry -ForegroundColor Red
        }
        else {
            Write-Host $logEntry
        }
    }
}

# Function to process and write logs from the queue to the log file
function Process-Logs {
    $logEntry = $null
    while ($logQueue.TryDequeue([ref]$logEntry)) {
        Add-Content -Path $LogFilePath -Value $logEntry.Value
    }
}

function Process-Errors {
    $errorEntry = $null
    while ($errorQueue.TryDequeue([ref]$errorEntry)) {
        Add-Content -Path $LogFilePath -Value $errorEntry.Value
    }
}

# Set default log file path if not provided
if (-not $LogFilePath) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogDirectory = Join-Path $LocalBackupRoot "Logs"
    if (!(Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }
    $LogFilePath = Join-Path $LogDirectory "BackupLog_$timestamp.txt"
}

# Start logging
Enqueue-Log "Backup script started." -Output

# Install PnP.PowerShell Module if not already installed
if (!(Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Enqueue-Log "Installing PnP.PowerShell module." -Output
    try {
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Enqueue-Log "PnP.PowerShell module installed successfully." -Output
    }
    catch {
        Enqueue-Log "Error installing PnP.PowerShell module: $_" -Level "ERROR" -Output
        throw
    }
}

# Import the PnP.PowerShell module
Import-Module PnP.PowerShell -ErrorAction Stop
Enqueue-Log "PnP.PowerShell module imported successfully." -Output

# Connect to SharePoint Online Site using the registered Entra ID application
try {
    Enqueue-Log "Connecting to SharePoint Online site: $SiteUrl" -Output
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive -ErrorAction Stop
    Enqueue-Log "Successfully connected to SharePoint Online." -Output
}
catch {
    Enqueue-Log "Error connecting to SharePoint Online: $_" -Level "ERROR" -Output
    Disconnect-PnPOnline
    Process-Logs
    exit 1
}

# Get the site object to use in URL conversions
try {
    $Site = Get-PnPWeb -ErrorAction Stop
    Enqueue-Log "Retrieved Site Object: $($Site.Title)" -Output
}
catch {
    Enqueue-Log "Error retrieving Site Object: $_" -Level "ERROR" -Output
    Disconnect-PnPOnline
    Process-Logs
    exit 1
}

# Function to download files recursively with underscore filtering and concurrency
function Download-Files {
    param (
        [string]$ServerRelativeUrl,
        [string]$LocalPath
    )
    Enqueue-Log "Processing: $ServerRelativeUrl" -Output

    # Try to get the folder
    try {
        $folder = Get-PnPFolder -ServerRelativeUrl $ServerRelativeUrl -ErrorAction Stop
        $isFolder = $true
    }
    catch {
        $isFolder = $false
    }

    if ($isFolder) {
        # It's a folder, check if it contains an underscore
        if ($folder.Name -notmatch '_') {
            Enqueue-Log "Skipped folder (no underscore): $ServerRelativeUrl" -Level "INFO"
            return
        }

        Enqueue-Log "Folder contains underscore: $($folder.Name)" -Output

        # Convert Server-Relative URL to Site-Relative URL
        if ($Site.ServerRelativeUrl -ne "/") {
            $SiteRelativeUrl = $ServerRelativeUrl.Substring($Site.ServerRelativeUrl.Length).TrimStart('/')
        }
        else {
            $SiteRelativeUrl = $ServerRelativeUrl.TrimStart('/')
        }

        # Get items in the folder
        try {
            $items = Get-PnPFolderItem -FolderSiteRelativeUrl $SiteRelativeUrl -ItemType All -ErrorAction Stop
            Enqueue-Log "Number of items retrieved: $($items.Count)" -Output
        }
        catch {
            Enqueue-Log "Error retrieving items from folder '$ServerRelativeUrl': $_" -Level "ERROR"
            return
        }

        # Separate folders and files
        $subFolders = @()
        $files = @()

        foreach ($item in $items) {
            if ($item.Folder -ne $null) {
                # It's a subfolder
                $subFolders += $item.Folder.ServerRelativeUrl
            }
            elseif ($item.File -ne $null) {
                # It's a file
                $files += $item.File.ServerRelativeUrl
            }
        }

        # Create local directory if it doesn't exist
        if (!(Test-Path $LocalPath)) {
            try {
                New-Item -ItemType Directory -Path $LocalPath -Force | Out-Null
                Enqueue-Log "Created local directory: $LocalPath" -Output
            }
            catch {
                Enqueue-Log "Error creating local directory '$LocalPath': $_" -Level "ERROR"
                return
            }
        }

        # Download files in parallel
        if ($files.Count -gt 0) {
            Enqueue-Log "Starting download of $($files.Count) files from $ServerRelativeUrl" -Output

            # Define a throttle limit for parallel downloads
            $ThrottleLimit = 20

            # Process files in parallel batches
            $files | ForEach-Object -Parallel {
                param ($fileUrl, $LocalPath, $SiteRelativeUrl, $SiteServerRelativeUrl, $logQueue, $errorQueue)

                try {
                    # Determine the relative path for the file
                    if ($SiteServerRelativeUrl -ne "/") {
                        $relativePath = $fileUrl.Substring($SiteServerRelativeUrl.Length).TrimStart('/')
                    }
                    else {
                        $relativePath = $fileUrl.TrimStart('/')
                    }

                    $fileName = Split-Path $relativePath -Leaf
                    $destinationPath = Join-Path $LocalPath $fileName

                    # Download the file
                    Get-PnPFile -Url $fileUrl -Path $LocalPath -FileName $fileName -AsFile -Force -ErrorAction Stop

                    # Log success
                    $logMessage = "Downloaded file: $fileUrl"
                    $logQueue.Enqueue($logMessage)
                }
                catch {
                    # Log error
                    $errorMessage = "Error downloading file '$fileUrl': $_"
                    $errorQueue.Enqueue($errorMessage)
                }
            } -ThrottleLimit $ThrottleLimit -ArgumentList $_, $LocalPath, $SiteRelativeUrl, $Site.ServerRelativeUrl, $logQueue, $errorQueue
        }

        # Process subfolders in parallel
        if ($subFolders.Count -gt 0) {
            Enqueue-Log "Starting processing of $($subFolders.Count) subfolders from $ServerRelativeUrl" -Output

            # Define a throttle limit for parallel folder processing
            $ThrottleLimitFolders = 10

            # Process subfolders in parallel
            $subFolders | ForEach-Object -Parallel {
                param ($subFolderUrl, $LocalPath, $logQueue, $errorQueue)

                try {
                    # Determine the subfolder name
                    $folderName = Split-Path $subFolderUrl -Leaf
                    $subFolderLocalPath = Join-Path $LocalPath $folderName

                    # Ensure the local subfolder exists
                    if (!(Test-Path $subFolderLocalPath)) {
                        New-Item -ItemType Directory -Path $subFolderLocalPath -Force | Out-Null
                        $logQueue.Enqueue("Created local subfolder: $subFolderLocalPath")
                    }

                    # Recurse into the subfolder
                    Download-Files -ServerRelativeUrl $subFolderUrl -LocalPath $subFolderLocalPath
                }
                catch {
                    # Log error
                    $errorMessage = "Error processing subfolder '$subFolderUrl': $_"
                    $errorQueue.Enqueue($errorMessage)
                }
            } -ThrottleLimit $ThrottleLimitFolders -ArgumentList $_, $LocalPath, $logQueue, $errorQueue
        }
    }
    else {
        # It might be a Document Library
        Enqueue-Log "Attempting to retrieve items from Document Library at '$ServerRelativeUrl'" -Output

        # Get the list (Document Library)
        try {
            $list = Get-PnPList | Where-Object { $_.RootFolder.ServerRelativeUrl -eq $ServerRelativeUrl } | Select-Object -First 1
            if ($null -eq $list) {
                throw "No list found at '$ServerRelativeUrl'."
            }
            Enqueue-Log "Found Document Library: $($list.Title)" -Output
        }
        catch {
            Enqueue-Log "Error retrieving Document Library at '$ServerRelativeUrl': $_" -Level "ERROR"
            return
        }

        # Get all items in the library
        try {
            $listItems = Get-PnPListItem -List $list -PageSize 1000 -Fields FileLeafRef, FileRef, FSObjType -ErrorAction Stop
            Enqueue-Log "Number of items retrieved from library '$($list.Title)': $($listItems.Count)" -Output
        }
        catch {
            Enqueue-Log "Error retrieving items from Document Library '$($list.Title)': $_" -Level "ERROR"
            return
        }

        # Separate folders and files
        $folders = @()
        $files = @()

        foreach ($item in $listItems) {
            $fileRef = $item["FileRef"]
            $fileLeafRef = $item["FileLeafRef"]
            $isFolder = $item["FSObjType"] -eq 1

            if ($isFolder) {
                # It's a folder, check if it contains an underscore
                $folderName = Split-Path $fileRef -Leaf
                if ($folderName -match '_') {
                    $folders += $fileRef
                }
                else {
                    Enqueue-Log "Skipped Document Library folder (no underscore): $fileRef" -Level "INFO"
                }
            }
            else {
                # It's a file
                $files += $fileRef
            }
        }

        # Create local directory if it doesn't exist
        if (!(Test-Path $LocalPath)) {
            try {
                New-Item -ItemType Directory -Path $LocalPath -Force | Out-Null
                Enqueue-Log "Created local directory: $LocalPath" -Output
            }
            catch {
                Enqueue-Log "Error creating local directory '$LocalPath': $_" -Level "ERROR"
                return
            }
        }

        # Download files in parallel
        if ($files.Count -gt 0) {
            Enqueue-Log "Starting download of $($files.Count) files from Document Library '$($list.Title)'" -Output

            # Define a throttle limit for parallel downloads
            $ThrottleLimit = 20

            # Process files in parallel
            $files | ForEach-Object -Parallel {
                param ($fileUrl, $LocalPath, $SiteRelativeUrl, $SiteServerRelativeUrl, $logQueue, $errorQueue)

                try {
                    # Determine the relative path for the file
                    if ($SiteServerRelativeUrl -ne "/") {
                        $relativePath = $fileUrl.Substring($SiteServerRelativeUrl.Length).TrimStart('/')
                    }
                    else {
                        $relativePath = $fileUrl.TrimStart('/')
                    }

                    $fileName = Split-Path $relativePath -Leaf
                    $destinationPath = Join-Path $LocalPath $fileName

                    # Download the file
                    Get-PnPFile -Url $fileUrl -Path $LocalPath -FileName $fileName -AsFile -Force -ErrorAction Stop

                    # Log success
                    $logMessage = "Downloaded file: $fileUrl"
                    $logQueue.Enqueue($logMessage)
                }
                catch {
                    # Log error
                    $errorMessage = "Error downloading file '$fileUrl': $_"
                    $errorQueue.Enqueue($errorMessage)
                }
            } -ThrottleLimit $ThrottleLimit -ArgumentList $_, $LocalPath, $SiteRelativeUrl, $Site.ServerRelativeUrl, $logQueue, $errorQueue
        }

        # Process folders in parallel
        if ($folders.Count -gt 0) {
            Enqueue-Log "Starting processing of $($folders.Count) folders from Document Library '$($list.Title)'" -Output

            # Define a throttle limit for parallel folder processing
            $ThrottleLimitFolders = 10

            # Process folders in parallel
            $folders | ForEach-Object -Parallel {
                param ($folderUrl, $LocalPath, $logQueue, $errorQueue)

                try {
                    # Determine the folder name
                    $folderName = Split-Path $folderUrl -Leaf
                    $folderLocalPath = Join-Path $LocalPath $folderName

                    # Ensure the local folder exists
                    if (!(Test-Path $folderLocalPath)) {
                        New-Item -ItemType Directory -Path $folderLocalPath -Force | Out-Null
                        $logQueue.Enqueue("Created local folder: $folderLocalPath")
                    }

                    # Recurse into the folder
                    Download-Files -ServerRelativeUrl $folderUrl -LocalPath $folderLocalPath
                }
                catch {
                    # Log error
                    $errorMessage = "Error processing folder '$folderUrl': $_"
                    $errorQueue.Enqueue($errorMessage)
                }
            } -ThrottleLimit $ThrottleLimitFolders -ArgumentList $_, $LocalPath, $logQueue, $errorQueue
        }
    }
}

# Main Execution
try {
    if ($SubFolderServerRelativeUrl) {
        # Start downloading from the specified subfolder or library
        Enqueue-Log "Backing up from: $SubFolderServerRelativeUrl" -Output

        # Ensure the path starts with '/'
        if (-not $SubFolderServerRelativeUrl.StartsWith('/')) {
            $SubFolderServerRelativeUrl = '/' + $SubFolderServerRelativeUrl
        }

        # Determine the local path for the backup
        $folderName = Split-Path $SubFolderServerRelativeUrl -Leaf
        $localBackupPath = Join-Path $LocalBackupRoot $folderName

        # Create local directory if it doesn't exist
        if (!(Test-Path $localBackupPath)) {
            New-Item -ItemType Directory -Path $localBackupPath -Force | Out-Null
            Enqueue-Log "Created local backup directory: $localBackupPath" -Output
        }

        # Start downloading
        Download-Files -ServerRelativeUrl $SubFolderServerRelativeUrl -LocalPath $localBackupPath
    }
    else {
        # Backup all Document Libraries
        try {
            Enqueue-Log "Retrieving document libraries." -Output
            $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }  # 101 is the template ID for Document Library
            Enqueue-Log "Number of document libraries retrieved: $($libraries.Count)" -Output
        }
        catch {
            Enqueue-Log "Error retrieving document libraries: $_" -Level "ERROR"
            Disconnect-PnPOnline
            Process-Logs
            exit 1
        }

        foreach ($library in $libraries) {
            $libraryTitle = $library.Title
            $localLibraryPath = Join-Path $LocalBackupRoot $libraryTitle

            # Create local directory if it doesn't exist
            if (!(Test-Path $localLibraryPath)) {
                try {
                    New-Item -ItemType Directory -Path $localLibraryPath -Force | Out-Null
                    Enqueue-Log "Created local library directory: $localLibraryPath" -Output
                }
                catch {
                    Enqueue-Log "Error creating local library directory '$localLibraryPath': $_" -Level "ERROR"
                    continue
                }
            }

            Enqueue-Log "Backing up library: $libraryTitle" -Output

            # Start downloading from the root folder of the library
            $rootFolderServerRelativeUrl = $library.RootFolder.ServerRelativeUrl
            Enqueue-Log "Root folder Server Relative URL: $rootFolderServerRelativeUrl" -Output

            # Proceed with the root folder (library root)
            Download-Files -ServerRelativeUrl $rootFolderServerRelativeUrl -LocalPath $localLibraryPath
        }
    }
}
catch {
    Enqueue-Log "An unexpected error occurred: $_" -Level "ERROR"
}
finally {
    # Process remaining logs
    Process-Logs
    Process-Errors

    # Disconnect from SharePoint Online
    try {
        Disconnect-PnPOnline -ErrorAction Stop
        Enqueue-Log "Disconnected from SharePoint Online." -Output
    }
    catch {
        Enqueue-Log "Error disconnecting from SharePoint Online: $_" -Level "ERROR"
    }

    # Finalize logging
    Enqueue-Log "Backup script completed." -Output
    Process-Logs

    # Display a final message
    Write-Host "Backup process finished. Check the log file at '$LogFilePath' for detailed information." -ForegroundColor Green
}
