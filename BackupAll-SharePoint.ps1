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

    [string]$LogFilePath
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

# Logging function with thread-safe append
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$Output
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    if ($Output) {
        Write-Host $logEntry
    }
    # Use a synchronized lock to ensure thread-safe logging
    $script:LogLock = $script:LogLock -or ([object]::new())
    [System.Threading.Monitor]::Enter($script:LogLock)
    try {
        $logEntry | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    }
    finally {
        [System.Threading.Monitor]::Exit($script:LogLock)
    }
}

# Start logging
Write-Log "Backup script started." -Output

# Install PnP.PowerShell Module if not already installed
if (!(Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Log "Installing PnP.PowerShell module." -Output
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}

# Connect to SharePoint Online Site using the registered Entra ID application
try {
    Write-Log "Connecting to SharePoint Online site: $SiteUrl" -Output
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive
    Write-Log "Successfully connected to SharePoint Online." -Output
}
catch {
    Write-Log "Error connecting to SharePoint Online: $_" "ERROR" -Output
    exit 1
}

# Get the site object to use in URL conversions
try {
    $Site = Get-PnPWeb
    Write-Log "Retrieved Site Object: $($Site.Title)" -Output
}
catch {
    Write-Log "Error retrieving Site Object: $_" "ERROR" -Output
    Disconnect-PnPOnline
    exit 1
}

# Function to establish a new PnP connection for parallel tasks
function Get-PnPConnection {
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -PnPManagementShell -ErrorAction Stop
        return $true
    }
    catch {
        Write-Log "Error establishing parallel PnP connection: $_" "ERROR" -Output
        return $false
    }
}

# Function to download files recursively with underscore filtering
function Download-Files {
    param (
        [string]$ServerRelativeUrl,
        [string]$LocalPath
    )
    Write-Log "Processing: $ServerRelativeUrl" -Output

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
            Write-Log "Skipped folder (no underscore): $ServerRelativeUrl" "INFO" -Output
            return
        }

        Write-Log "Folder contains underscore: $($folder.Name)"

        # Convert Server-Relative URL to Site-Relative URL
        if ($Site.ServerRelativeUrl -ne "/") {
            $SiteRelativeUrl = $ServerRelativeUrl.Substring($Site.ServerRelativeUrl.Length).TrimStart('/')
        }
        else {
            $SiteRelativeUrl = $ServerRelativeUrl.TrimStart('/')
        }

        # Get items in the folder
        try {
            $items = Get-PnPFolderItem -FolderSiteRelativeUrl $SiteRelativeUrl -ItemType All
            Write-Log "Number of items retrieved: $($items.Count)" -Output
        }
        catch {
            Write-Log "Error retrieving items from folder '$ServerRelativeUrl': $_" "ERROR" -Output
            return
        }

        # Process items in parallel
        $items | ForEach-Object -Parallel {
            param ($item, $ServerRelativeUrl, $LocalPath, $SiteUrl, $ClientId, $TenantId, $LogFilePath)

            # Establish a new PnP connection in the parallel runspace
            if (-not (Get-PnPConnection)) {
                Write-Log "Failed to establish PnP connection for parallel task." "ERROR" -Output
                return
            }

            # Redefine Write-Log for parallel tasks
            function Write-LogParallel {
                param (
                    [string]$Message,
                    [string]$Level = "INFO"
                )
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $logEntry = "[$timestamp] [$Level] $Message"
                Write-Host $logEntry
                # Use a synchronized lock to ensure thread-safe logging
                $script:LogLock = $script:LogLock -or ([object]::new())
                [System.Threading.Monitor]::Enter($script:LogLock)
                try {
                    $logEntry | Out-File -FilePath $using:LogFilePath -Append -Encoding UTF8
                }
                finally {
                    [System.Threading.Monitor]::Exit($script:LogLock)
                }
            }

            if ($item.Folder -ne $null) {
                # It's a subfolder
                $folderName = $item.Name

                # Check if subfolder contains an underscore
                if ($folderName -notmatch '_') {
                    Write-LogParallel "Skipped subfolder (no underscore): $($item.Folder.ServerRelativeUrl)" "INFO"
                    return
                }

                $subFolderServerRelativeUrl = $item.Folder.ServerRelativeUrl
                $subFolderPath = Join-Path $LocalPath $folderName

                Write-LogParallel "Processing subfolder with underscore: $subFolderServerRelativeUrl"

                if (!(Test-Path $subFolderPath)) {
                    New-Item -ItemType Directory -Path $subFolderPath | Out-Null
                }

                # Recurse into subfolder
                Download-Files -ServerRelativeUrl $subFolderServerRelativeUrl -LocalPath $subFolderPath
            }
            elseif ($item.File -ne $null) {
                $fileName = $item.Name
                Write-LogParallel "Processing file with underscore: $($item.File.ServerRelativeUrl)"
                try {
                    # Download the file
                    Get-PnPFile -Url $item.File.ServerRelativeUrl -Path $LocalPath -FileName $fileName -AsFile -Force
                    Write-LogParallel "Downloaded file: $($item.File.ServerRelativeUrl)"
                }
                catch {
                    Write-LogParallel "Error downloading file '$($item.File.ServerRelativeUrl)': $_" "ERROR"
                }
            }

            # Disconnect the parallel PnP connection
            Disconnect-PnPOnline
        } -ThrottleLimit 10
    }
    else {
        # It might be a Document Library
        Write-Log "Attempting to retrieve items from Document Library at '$ServerRelativeUrl'" -Output

        # Get the list (Document Library)
        try {
            $list = Get-PnPList | Where-Object { $_.RootFolder.ServerRelativeUrl -eq $ServerRelativeUrl } | Select-Object -First 1
            if ($null -eq $list) {
                throw "No list found at '$ServerRelativeUrl'"
            }
            Write-Log "Found Document Library: $($list.Title)" -Output
        }
        catch {
            Write-Log "Error retrieving Document Library at '$ServerRelativeUrl': $_" "ERROR" -Output
            return
        }

        # Get all items in the library
        try {
            $listItems = Get-PnPListItem -List $list -PageSize 1000 -Fields FileLeafRef,FileRef,FSObjType
            Write-Log "Number of items retrieved from library: $($listItems.Count)" -Output
        }
        catch {
            Write-Log "Error retrieving items from Document Library '$($list.Title)': $_" "ERROR" -Output
            return
        }

        # Process list items in parallel
        $listItems | ForEach-Object -Parallel {
            param ($item, $ServerRelativeUrl, $LocalPath, $SiteUrl, $ClientId, $TenantId, $LogFilePath)

            # Establish a new PnP connection in the parallel runspace
            if (-not (Get-PnPConnection)) {
                Write-Log "Failed to establish PnP connection for parallel task." "ERROR" -Output
                return
            }

            # Redefine Write-Log for parallel tasks
            function Write-LogParallel {
                param (
                    [string]$Message,
                    [string]$Level = "INFO"
                )
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $logEntry = "[$timestamp] [$Level] $Message"
                Write-Host $logEntry
                # Use a synchronized lock to ensure thread-safe logging
                $script:LogLock = $script:LogLock -or ([object]::new())
                [System.Threading.Monitor]::Enter($script:LogLock)
                try {
                    $logEntry | Out-File -FilePath $using:LogFilePath -Append -Encoding UTF8
                }
                finally {
                    [System.Threading.Monitor]::Exit($script:LogLock)
                }
            }

            $fileRef = $item["FileRef"]
            $fileLeafRef = $item["FileLeafRef"]
            $isFolder = $item["FSObjType"] -eq 1

            $relativePath = if ($using:Site.ServerRelativeUrl -ne "/") {
                $fileRef.Substring($using:Site.ServerRelativeUrl.Length).TrimStart('/')
            }
            else {
                $fileRef.TrimStart('/')
            }
            $localItemPath = Join-Path $LocalPath $relativePath

            if ($isFolder) {
                # It's a folder, check if it contains an underscore
                $folderName = Split-Path $relativePath -Leaf
                if ($folderName -notmatch '_') {
                    Write-LogParallel "Skipped doclib (no underscore): $fileRef" "INFO"
                    return
                }

                Write-LogParallel "Processing folder with underscore: $fileRef"

                if (!(Test-Path $localItemPath)) {
                    New-Item -ItemType Directory -Path $localItemPath | Out-Null
                }
            }
            else {
                Write-LogParallel "Processing file with underscore: $fileRef"
                $localFolderPath = Split-Path $localItemPath -Parent
                if (!(Test-Path $localFolderPath)) {
                    New-Item -ItemType Directory -Path $localFolderPath -Force | Out-Null
                }

                try {
                    # Download the file
                    Get-PnPFile -Url $fileRef -Path $localFolderPath -FileName $fileLeafRef -AsFile -Force
                    Write-LogParallel "Downloaded file: $fileRef"
                }
                catch {
                    Write-LogParallel "Error downloading file '$fileRef': $_" "ERROR"
                }
            }

            # Disconnect the parallel PnP connection
            Disconnect-PnPOnline
        } -ThrottleLimit 10
    }
}

# Main Execution
if ($SubFolderServerRelativeUrl) {
    # Start downloading from the specified subfolder or library
    Write-Log "Backing up from: $SubFolderServerRelativeUrl" -Output

    # Ensure the path starts with '/'
    if (-not $SubFolderServerRelativeUrl.StartsWith('/')) {
        $SubFolderServerRelativeUrl = '/' + $SubFolderServerRelativeUrl
    }

    # Start downloading
    Download-Files -ServerRelativeUrl $SubFolderServerRelativeUrl -LocalPath $LocalBackupRoot
}
else {
    # Backup all Document Libraries
    try {
        Write-Log "Retrieving document libraries." -Output
        $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }  # 101 is the template ID for Document Library
    }
    catch {
        Write-Log "Error retrieving document libraries: $_" "ERROR" -Output
        Disconnect-PnPOnline
        exit 1
    }

    # Process each library in parallel
    $libraries | ForEach-Object -Parallel {
        param ($library, $LocalBackupRoot, $SiteUrl, $ClientId, $TenantId, $LogFilePath)

        # Establish a new PnP connection in the parallel runspace
        if (-not (Get-PnPConnection)) {
            Write-Log "Failed to establish PnP connection for parallel task." "ERROR"
            return
        }

        # Redefine Write-Log for parallel tasks
        function Write-LogParallel {
            param (
                [string]$Message,
                [string]$Level = "INFO"
            )
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "[$timestamp] [$Level] $Message"
            Write-Host $logEntry
            # Use a synchronized lock to ensure thread-safe logging
            $script:LogLock = $script:LogLock -or ([object]::new())
            [System.Threading.Monitor]::Enter($script:LogLock)
            try {
                $logEntry | Out-File -FilePath $using:LogFilePath -Append -Encoding UTF8
            }
            finally {
                [System.Threading.Monitor]::Exit($script:LogLock)
            }
        }

        $libraryTitle = $library.Title
        $localLibraryPath = Join-Path $using:LocalBackupRoot $libraryTitle

        # Create local directory if it doesn't exist
        if (!(Test-Path $localLibraryPath)) {
            New-Item -ItemType Directory -Path $localLibraryPath | Out-Null
        }

        Write-LogParallel "Backing up library: $libraryTitle"

        # Start downloading from the root folder of the library
        $rootFolderServerRelativeUrl = $library.RootFolder.ServerRelativeUrl
        Write-LogParallel "Root folder Server Relative URL: $rootFolderServerRelativeUrl"

        # Proceed with the root folder (library root)
        Download-Files -ServerRelativeUrl $rootFolderServerRelativeUrl -LocalPath $localLibraryPath

        # Disconnect the parallel PnP connection
        Disconnect-PnPOnline
    } -ThrottleLimit 10 -ArgumentList $_, $LocalBackupRoot, $SiteUrl, $ClientId, $TenantId, $LogFilePath
}

# Disconnect from SharePoint Online
Disconnect-PnPOnline
Write-Log "Backup script completed."
