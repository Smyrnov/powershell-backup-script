# Requires PowerShell 5.1 or later

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
    [string]$SubFolderServerRelativeUrl,  # Server-relative URL of the library or folder

    [Parameter(Mandatory = $false)]
    $LastModifiedDate,          # Optional parameter for filtering based on last modified date

    [Parameter(Mandatory = $false)]
    $CreatedDate,               # Optional parameter for filtering based on created date

    [string]$LogFilePath
)

# Ensure PowerShell 5.1 or later is being used
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "This script requires PowerShell version 5.1 or later."
    exit 1
}

# Import necessary .NET namespaces
Add-Type -AssemblyName System.Collections.Concurrent

# Initialize concurrent queues for logs and errors
$logQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$errorQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# Function to enqueue logs
function Enqueue-Log {
    param (
        [string]$Message
    )
    $logQueue.Enqueue($Message)
}

# Function to enqueue errors
function Enqueue-ErrorLog {
    param (
        [string]$Message
    )
    $errorQueue.Enqueue($Message)
}

# Set default log file path if not provided
if (-not $LogFilePath) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogDirectory = Join-Path $LocalBackupRoot "Logs"
    if (!(Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory | Out-Null
    }
    $LogFilePath = Join-Path $LogDirectory "BackupLog_$timestamp.txt"
}

# Start logging
Enqueue-Log "Backup script started."

# Install PnP.PowerShell Module if not already installed
if (!(Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Enqueue-Log "Installing PnP.PowerShell module."
    try {
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Enqueue-Log "PnP.PowerShell module installed successfully."
    }
    catch {
        Enqueue-ErrorLog "Error installing PnP.PowerShell module: $_"
        # Exit if module installation fails
        Enqueue-ErrorLog "Backup script terminated due to module installation failure."
        Write-Host "Backup process finished with errors. Check the log file at '$LogFilePath' for detailed information." -ForegroundColor Red
        exit 1
    }
}

# Connect to SharePoint Online Site using the registered Entra ID application
try {
    Enqueue-Log "Connecting to SharePoint Online site: $SiteUrl"
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive -ErrorAction Stop
    Enqueue-Log "Successfully connected to SharePoint Online."
}
catch {
    Enqueue-ErrorLog "Error connecting to SharePoint Online: $_"
    Enqueue-ErrorLog "Backup script terminated due to connection failure."
    Write-Host "Backup process finished with errors. Check the log file at '$LogFilePath' for detailed information." -ForegroundColor Red
    exit 1
}

# Get the site object to use in URL conversions
try {
    $Site = Get-PnPWeb -ErrorAction Stop
}
catch {
    Enqueue-ErrorLog "Error retrieving site information: $_"
    Enqueue-ErrorLog "Backup script terminated due to site retrieval failure."
    Write-Host "Backup process finished with errors. Check the log file at '$LogFilePath' for detailed information." -ForegroundColor Red
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    exit 1
}

# Function to download files recursively with underscore filtering and parallel processing
function Download-Files {
    param (
        [string]$ServerRelativeUrl,
        [string]$LocalPath,
        $LastModifiedDate,
        $CreatedDate
    )

    Enqueue-Log "Processing: $ServerRelativeUrl"

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
            Enqueue-Log "Skipped folder (no underscore): $ServerRelativeUrl"
            return
        }

        Enqueue-Log "Folder contains underscore: $($folder.Name)"

        # Convert Server-Relative URL to Site-Relative URL
        $SiteRelativeUrl = $ServerRelativeUrl.Substring($Site.ServerRelativeUrl.Length).TrimStart('/')

        # Get items in the folder
        try {
            $items = Get-PnPFolderItem -FolderSiteRelativeUrl $SiteRelativeUrl -ItemType All -ErrorAction Stop
            Enqueue-Log "Number of items retrieved: $($items.Count)"
        }
        catch {
            Enqueue-ErrorLog "Error retrieving items from folder '$ServerRelativeUrl': $_"
            return
        }

        # Prepare lists for parallel processing
        $foldersToProcess = @()
        $filesToDownload = @()

        foreach ($item in $items) {
            if ($item.Folder -ne $null) {
                # It's a subfolder
                $folderName = $item.Name
                $subFolderServerRelativeUrl = $item.Folder.ServerRelativeUrl

                if ($folderName -match '_') {
                    $foldersToProcess += $subFolderServerRelativeUrl
                }
                else {
                    Enqueue-Log "Skipped subfolder (no underscore): $subFolderServerRelativeUrl"
                }
            }
            elseif ($item.File -ne $null) {
                # It's a file (no filtering on file names)
                $fileServerRelativeUrl = $item.File.ServerRelativeUrl
                $filesToDownload += $fileServerRelativeUrl
            }
        }

        # Download files in parallel
        if ($filesToDownload.Count -gt 0) {
            Enqueue-Log "Starting parallel download of $($filesToDownload.Count) files from $ServerRelativeUrl"

            $jobs = @()

            foreach ($fileUrl in $filesToDownload) {
                $jobs += Start-Job -ScriptBlock {
                    param($fileUrl, $LocalPath, $LastModifiedDate, $CreatedDate, $logQueue, $errorQueue)

                    try {
                        # Determine the local file path
                        $fileName = Split-Path $fileUrl -Leaf
                        $destinationPath = Join-Path $LocalPath $fileName

                        $downloadFile = $true

                        if ($LastModifiedDate -or $CreatedDate) {
                            if (Test-Path $destinationPath) {
                                # Get the SharePoint file's last modified and creation dates
                                $spFile = Get-PnPFile -Url $fileUrl -AsListItem -ErrorAction Stop
                                $spLastModified = [datetime]$spFile["Modified"]
                                $spCreated = [datetime]$spFile["Created"]

                                # Determine if file should be downloaded
                                $shouldDownload = $false

                                if ($LastModifiedDate) {
                                    if ($spLastModified -gt $LastModifiedDate) {
                                        $shouldDownload = $true
                                    }
                                }

                                if ($CreatedDate) {
                                    if ($spCreated -gt $CreatedDate) {
                                        $shouldDownload = $true
                                    }
                                }

                                if (-not $shouldDownload) {
                                    # Skip downloading
                                    $logQueue.Enqueue("Skipped file (not created/modified after specified dates): $fileUrl")
                                    $downloadFile = $false
                                }
                            }
                            else {
                                # File does not exist locally; download it
                                $downloadFile = $true
                            }
                        }

                        if ($downloadFile) {
                            # Ensure the local directory exists
                            if (!(Test-Path $LocalPath)) {
                                New-Item -ItemType Directory -Path $LocalPath -Force | Out-Null
                            }

                            # Download the file
                            Get-PnPFile -Url $fileUrl -Path $LocalPath -FileName $fileName -AsFile -Force -ErrorAction Stop
                            $logQueue.Enqueue("Downloaded file: $fileUrl")
                        }
                    }
                    catch {
                        $errorQueue.Enqueue("Error downloading file '$fileUrl': $_")
                    }
                } -ArgumentList $fileUrl, $LocalPath, $LastModifiedDate, $CreatedDate, $logQueue, $errorQueue
            }

            # Wait for all jobs to complete
            $jobs | Wait-Job | Receive-Job

            # Remove completed jobs
            $jobs | Remove-Job
        }

        # Process subfolders recursively
        foreach ($subFolder in $foldersToProcess) {
            # Determine the local path for the subfolder
            $folderName = Split-Path $subFolder -Leaf
            $subFolderLocalPath = Join-Path $LocalPath $folderName

            if (!(Test-Path $subFolderLocalPath)) {
                New-Item -ItemType Directory -Path $subFolderLocalPath | Out-Null
            }

            # Recurse into the subfolder
            Download-Files -ServerRelativeUrl $subFolder -LocalPath $subFolderLocalPath -LastModifiedDate $LastModifiedDate -CreatedDate $CreatedDate
        }
    }

    # Main Execution Block with Structured Error Handling
    try {
        if ($SubFolderServerRelativeUrl) {
            # Start downloading from the specified subfolder or library
            Enqueue-Log "Backing up from: $SubFolderServerRelativeUrl"

            # Ensure the path starts with '/'
            if (-not $SubFolderServerRelativeUrl.StartsWith('/')) {
                $SubFolderServerRelativeUrl = '/' + $SubFolderServerRelativeUrl
            }

            # Start downloading
            Download-Files -ServerRelativeUrl $SubFolderServerRelativeUrl -LocalPath $LocalBackupRoot -LastModifiedDate $LastModifiedDate -CreatedDate $CreatedDate
        }
        else {
            # Retrieve All Document Libraries
            try {
                Enqueue-Log "Retrieving document libraries."
                $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }  # 101 is the template ID for Document Library
                Enqueue-Log "Number of document libraries retrieved: $($libraries.Count)"
            }
            catch {
                Enqueue-ErrorLog "Error retrieving document libraries: $_"
                throw
            }

            foreach ($library in $libraries) {
                $libraryTitle = $library.Title
                $localLibraryPath = Join-Path $LocalBackupRoot $libraryTitle

                # Check if library name contains an underscore
                if ($libraryTitle -notmatch '_') {
                    Enqueue-Log "Skipped Document Library (no underscore): $libraryTitle"
                    continue
                }

                # Create local directory if it doesn't exist
                if (!(Test-Path $localLibraryPath)) {
                    New-Item -ItemType Directory -Path $localLibraryPath | Out-Null
                }

                Enqueue-Log "Backing up library: $libraryTitle"

                # Start downloading from the root folder of the library
                $rootFolderServerRelativeUrl = $library.RootFolder.ServerRelativeUrl
                Enqueue-Log "Root folder Server Relative URL: $rootFolderServerRelativeUrl"

                # Proceed with the root folder (library root)
                Download-Files -ServerRelativeUrl $rootFolderServerRelativeUrl -LocalPath $localLibraryPath -LastModifiedDate $LastModifiedDate -CreatedDate $CreatedDate
            }
        }
    }
    catch {
        Enqueue-ErrorLog "Backup script terminated due to an unexpected error: $_"
    }
    finally {
        # Finalize logging

        # Write first log message to the console
        if ($logQueue.Count -gt 0) {
            if ($logQueue.TryDequeue([ref]$firstLog)) {
                Write-Host $firstLog
            }
        }

        # Write all remaining logs to the log file
        while ($logQueue.TryDequeue([ref]$log)) {
            Add-Content -Path $LogFilePath -Value $log
        }

        # Write errors to the console
        if ($errorQueue.Count -gt 0) {
            Write-Host "Errors encountered during backup:" -ForegroundColor Red
            while ($errorQueue.TryDequeue([ref]$errorMsg)) {
                Write-Host $errorMsg -ForegroundColor Red
            }
        }

        # Disconnect from SharePoint Online
        try {
            Disconnect-PnPOnline -ErrorAction Stop
            Enqueue-Log "Disconnected from SharePoint Online."
        }
        catch {
            Enqueue-ErrorLog "Error disconnecting from SharePoint Online: $_"
        }

        # Write completion log
        Enqueue-Log "Backup script completed."

        # Write any remaining logs to the log file
        while ($logQueue.TryDequeue([ref]$log)) {
            Add-Content -Path $LogFilePath -Value $log
        }

        # Display a final message
        Write-Host "Backup process finished. Check the log file at '$LogFilePath' for detailed information." -ForegroundColor Green
    }
}
