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

    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,  # New parameter for Client Secret

    [Parameter(Mandatory = $false)]
    [int]$DegreeOfParallelism = 5  # New parameter to control concurrency
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

# Initialize a lock object for thread-safe logging
$Script:LogLock = New-Object Object

# Logging function with thread-safe append
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
    # Use a synchronized lock to ensure thread-safe logging
    [System.Threading.Monitor]::Enter($Script:LogLock)
    try {
        $logEntry | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    }
    finally {
        [System.Threading.Monitor]::Exit($Script:LogLock)
    }
}

# Start logging
Write-Log "Backup script started."

# Install PnP.PowerShell Module if not already installed
if (!(Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Log "Installing PnP.PowerShell module."
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}

# Function to establish a new PnP connection for parallel tasks
function Establish-PnPConnection {
    param (
        [string]$SiteUrl,
        [string]$ClientId,
        [string]$TenantId,
        [string]$ClientSecret
    )
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -ClientSecret $ClientSecret -ErrorAction Stop
        Write-Log "Established PnP connection to $SiteUrl."
        return $true
    }
    catch {
        Write-Log "Error establishing PnP connection: $_" "ERROR"
        return $false
    }
}

# Connect to SharePoint Online Site using Client Secret
try {
    Write-Log "Connecting to SharePoint Online site: $SiteUrl"
    Establish-PnPConnection -SiteUrl $SiteUrl -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret
    Write-Log "Successfully connected to SharePoint Online."
}
catch {
    Write-Log "Error connecting to SharePoint Online: $_" "ERROR"
    exit 1
}

# Get the site object to use in URL conversions
try {
    $Site = Get-PnPWeb -ErrorAction Stop
    Write-Log "Retrieved Site Object: $($Site.Title)"
}
catch {
    Write-Log "Error retrieving Site Object: $_" "ERROR"
    Disconnect-PnPOnline
    exit 1
}

# Function to download files recursively with underscore filtering
function Download-Files {
    param (
        [string]$ServerRelativeUrl,
        [string]$LocalPath,
        [string]$SiteUrl,
        [string]$ClientId,
        [string]$TenantId,
        [string]$ClientSecret,
        [string]$LogFilePath,
        [int]$DegreeOfParallelism
    )

    # Each parallel runspace establishes its own connection
    $connected = $false
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -ClientSecret $ClientSecret -ErrorAction Stop
        $connected = $true
        Write-Log "Established parallel PnP connection to $SiteUrl."
    }
    catch {
        Write-Log "Error establishing parallel PnP connection: $_" "ERROR"
        return
    }

    Write-Log "Processing: $ServerRelativeUrl"

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
            Write-Log "Skipped folder (no underscore): $ServerRelativeUrl" "INFO"
            if ($connected) { Disconnect-PnPOnline }
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
            $items = Get-PnPFolderItem -FolderSiteRelativeUrl $SiteRelativeUrl -ItemType All -ErrorAction Stop
            Write-Log "Number of items retrieved: $($items.Count)"
        }
        catch {
            Write-Log "Error retrieving items from folder '$ServerRelativeUrl': $_" "ERROR"
            if ($connected) { Disconnect-PnPOnline }
            return
        }

        # Process items in parallel
        $items | ForEach-Object -Parallel {
            param (
                $item,
                $ServerRelativeUrl,
                $LocalPath,
                $SiteUrl,
                $ClientId,
                $TenantId,
                $ClientSecret,
                $LogFilePath,
                $ScriptLock
            )

            # Define thread-safe logging within parallel runspace
            function Write-LogParallel {
                param (
                    [string]$Message,
                    [string]$Level = "INFO"
                )
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $logEntry = "[$timestamp] [$Level] $Message"
                Write-Host $logEntry
                # Use the lock to synchronize log writing
                [System.Threading.Monitor]::Enter($ScriptLock.Value)
                try {
                    $logEntry | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
                }
                finally {
                    [System.Threading.Monitor]::Exit($ScriptLock.Value)
                }
            }

            # Establish a new PnP connection in the parallel runspace
            $connectedParallel = $false
            try {
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -ClientSecret $ClientSecret -ErrorAction Stop
                $connectedParallel = $true
                Write-LogParallel "Established parallel PnP connection to $SiteUrl."
            }
            catch {
                Write-LogParallel "Error establishing parallel PnP connection: $_" "ERROR"
                return
            }

            if ($item.Folder -ne $null) {
                # It's a subfolder
                $folderName = $item.Name

                # Check if subfolder contains an underscore
                if ($folderName -notmatch '_') {
                    Write-LogParallel "Skipped subfolder (no underscore): $($item.Folder.ServerRelativeUrl)" "INFO"
                    if ($connectedParallel) { Disconnect-PnPOnline }
                    return
                }

                $subFolderServerRelativeUrl = $item.Folder.ServerRelativeUrl
                $subFolderPath = Join-Path $LocalPath $folderName

                Write-LogParallel "Processing subfolder with underscore: $subFolderServerRelativeUrl"

                if (!(Test-Path $subFolderPath)) {
                    try {
                        New-Item -ItemType Directory -Path $subFolderPath -ErrorAction Stop | Out-Null
                        Write-LogParallel "Created local directory: $subFolderPath"
                    }
                    catch {
                        Write-LogParallel "Error creating directory '$subFolderPath': $_" "ERROR"
                        if ($connectedParallel) { Disconnect-PnPOnline }
                        return
                    }
                }

                # Recurse into subfolder
                Download-Files -ServerRelativeUrl $subFolderServerRelativeUrl -LocalPath $subFolderPath -SiteUrl $SiteUrl -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret -LogFilePath $LogFilePath -DegreeOfParallelism $using:DegreeOfParallelism
            }
            elseif ($item.File -ne $null) {
                $fileName = $item.Name
                Write-LogParallel "Processing file with underscore: $($item.File.ServerRelativeUrl)"
                try {
                    # Download the file
                    Get-PnPFile -Url $item.File.ServerRelativeUrl -Path $LocalPath -FileName $fileName -AsFile -Force -ErrorAction Stop
                    Write-LogParallel "Downloaded file: $($item.File.ServerRelativeUrl)"
                }
                catch {
                    Write-LogParallel "Error downloading file '$($item.File.ServerRelativeUrl)': $_" "ERROR"
                }
            }

            # Disconnect the parallel PnP connection
            try {
                Disconnect-PnPOnline -ErrorAction Stop
                Write-LogParallel "Disconnected parallel PnP connection."
            }
            catch {
                Write-LogParallel "Error disconnecting parallel PnP connection: $_" "ERROR"
            }
        } -ThrottleLimit $DegreeOfParallelism -ArgumentList $_, $ServerRelativeUrl, $LocalPath, $SiteUrl, $ClientId, $TenantId, $ClientSecret, $LogFilePath, $Script:LogLock

    }
    else {
        # It might be a Document Library
        Write-Log "Attempting to retrieve items from Document Library at '$ServerRelativeUrl'"

        # Get the list (Document Library)
        try {
            $list = Get-PnPList | Where-Object { $_.RootFolder.ServerRelativeUrl -eq $ServerRelativeUrl } | Select-Object -First 1
            if ($null -eq $list) {
                throw "No list found at '$ServerRelativeUrl'"
            }
            Write-Log "Found Document Library: $($list.Title)"
        }
        catch {
            Write-Log "Error retrieving Document Library at '$ServerRelativeUrl': $_" "ERROR"
            if ($connected) { Disconnect-PnPOnline }
            return
        }

        # Get all items in the library
        try {
            $listItems = Get-PnPListItem -List $list -PageSize 1000 -Fields FileLeafRef,FileRef,FSObjType -ErrorAction Stop
            Write-Log "Number of items retrieved from library: $($listItems.Count)"
        }
        catch {
            Write-Log "Error retrieving items from Document Library '$($list.Title)': $_" "ERROR"
            if ($connected) { Disconnect-PnPOnline }
            return
        }

        # Process list items in parallel
        $listItems | ForEach-Object -Parallel {
            param (
                $item,
                $ServerRelativeUrl,
                $LocalPath,
                $SiteUrl,
                $ClientId,
                $TenantId,
                $ClientSecret,
                $LogFilePath,
                $ScriptLock
            )

            # Define thread-safe logging within parallel runspace
            function Write-LogParallel {
                param (
                    [string]$Message,
                    [string]$Level = "INFO"
                )
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $logEntry = "[$timestamp] [$Level] $Message"
                Write-Host $logEntry
                # Use the lock to synchronize log writing
                [System.Threading.Monitor]::Enter($ScriptLock.Value)
                try {
                    $logEntry | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
                }
                finally {
                    [System.Threading.Monitor]::Exit($ScriptLock.Value)
                }
            }

            # Establish a new PnP connection in the parallel runspace
            $connectedParallel = $false
            try {
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -ClientSecret $ClientSecret -ErrorAction Stop
                $connectedParallel = $true
                Write-LogParallel "Established parallel PnP connection to $SiteUrl."
            }
            catch {
                Write-LogParallel "Error establishing parallel PnP connection: $_" "ERROR"
                return
            }

            $fileRef = $item["FileRef"]
            $fileLeafRef = $item["FileLeafRef"]
            $isFolder = $item["FSObjType"] -eq 1

            if (-not $fileRef) {
                Write-LogParallel "FileRef is null for item ID: $($item.Id)" "ERROR"
                if ($connectedParallel) { Disconnect-PnPOnline }
                return
            }

            $relativePath = if ($Site.ServerRelativeUrl -ne "/") {
                $fileRef.Substring($Site.ServerRelativeUrl.Length).TrimStart('/')
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
                    if ($connectedParallel) { Disconnect-PnPOnline }
                    return
                }

                Write-LogParallel "Processing folder with underscore: $fileRef"

                if (!(Test-Path $localItemPath)) {
                    try {
                        New-Item -ItemType Directory -Path $localItemPath -Force -ErrorAction Stop | Out-Null
                        Write-LogParallel "Created local directory: $localItemPath"
                    }
                    catch {
                        Write-LogParallel "Error creating directory '$localItemPath': $_" "ERROR"
                        if ($connectedParallel) { Disconnect-PnPOnline }
                        return
                    }
                }
            }
            else {
                Write-LogParallel "Processing file with underscore: $fileRef"
                $localFolderPath = Split-Path $localItemPath -Parent
                if (!(Test-Path $localFolderPath)) {
                    try {
                        New-Item -ItemType Directory -Path $localFolderPath -Force -ErrorAction Stop | Out-Null
                        Write-LogParallel "Created local directory: $localFolderPath"
                    }
                    catch {
                        Write-LogParallel "Error creating directory '$localFolderPath': $_" "ERROR"
                        if ($connectedParallel) { Disconnect-PnPOnline }
                        return
                    }
                }

                try {
                    # Download the file
                    Get-PnPFile -Url $fileRef -Path $localFolderPath -FileName $fileLeafRef -AsFile -Force -ErrorAction Stop
                    Write-LogParallel "Downloaded file: $fileRef"
                }
                catch {
                    Write-LogParallel "Error downloading file '$fileRef': $_" "ERROR"
                }
            }

            # Disconnect the parallel PnP connection
            try {
                Disconnect-PnPOnline -ErrorAction Stop
                Write-LogParallel "Disconnected parallel PnP connection."
            }
            catch {
                Write-LogParallel "Error disconnecting parallel PnP connection: $_" "ERROR"
            }
        } -ThrottleLimit $DegreeOfParallelism -ArgumentList $_, $ServerRelativeUrl, $LocalPath, $SiteUrl, $ClientId, $TenantId, $ClientSecret, $LogFilePath, $Script:LogLock
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
        Download-Files -ServerRelativeUrl $SubFolderServerRelativeUrl -LocalPath $LocalBackupRoot -SiteUrl $SiteUrl -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret -LogFilePath $LogFilePath -DegreeOfParallelism $DegreeOfParallelism
    }
    else {
        # Backup all Document Libraries
        try {
            Write-Log "Retrieving document libraries."
            $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }  # 101 is the template ID for Document Library
        }
        catch {
            Write-Log "Error retrieving document libraries: $_" "ERROR"
            if ($connected) { Disconnect-PnPOnline }
            exit 1
        }

        # Process each library in parallel
        $libraries | ForEach-Object -Parallel {
            param (
                $library,
                $LocalBackupRoot,
                $DegreeOfParallelism,
                $SiteUrl,
                $ClientId,
                $TenantId,
                $ClientSecret,
                $LogFilePath,
                $ScriptLock
            )

            # Define thread-safe logging within parallel runspace
            function Write-LogParallel {
                param (
                    [string]$Message,
                    [string]$Level = "INFO"
                )
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $logEntry = "[$timestamp] [$Level] $Message"
                Write-Host $logEntry
                # Use the lock to synchronize log writing
                [System.Threading.Monitor]::Enter($ScriptLock.Value)
                try {
                    $logEntry | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
                }
                finally {
                    [System.Threading.Monitor]::Exit($ScriptLock.Value)
                }
            }

            # Establish a new PnP connection in the parallel runspace
            $connectedParallel = $false
            try {
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -ClientSecret $ClientSecret -ErrorAction Stop
                $connectedParallel = $true
                Write-LogParallel "Established parallel PnP connection to $SiteUrl."
            }
            catch {
                Write-LogParallel "Error establishing parallel PnP connection: $_" "ERROR"
                return
            }

            $libraryTitle = $library.Title
            $localLibraryPath = Join-Path $LocalBackupRoot $libraryTitle

            # Create local directory if it doesn't exist
            if (!(Test-Path $localLibraryPath)) {
                try {
                    New-Item -ItemType Directory -Path $localLibraryPath -ErrorAction Stop | Out-Null
                    Write-LogParallel "Created local directory: $localLibraryPath"
                }
                catch {
                    Write-LogParallel "Error creating directory '$localLibraryPath': $_" "ERROR"
                    if ($connectedParallel) { Disconnect-PnPOnline }
                    return
                }
            }

            Write-LogParallel "Backing up library: $libraryTitle"

            # Start downloading from the root folder of the library
            $rootFolderServerRelativeUrl = $library.RootFolder.ServerRelativeUrl
            Write-LogParallel "Root folder Server Relative URL: $rootFolderServerRelativeUrl"

            # Proceed with the root folder (library root)
            Download-Files -ServerRelativeUrl $rootFolderServerRelativeUrl -LocalPath $localLibraryPath -SiteUrl $SiteUrl -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret -LogFilePath $LogFilePath -DegreeOfParallelism $DegreeOfParallelism

            # Disconnect the parallel PnP connection
            try {
                Disconnect-PnPOnline -ErrorAction Stop
                Write-LogParallel "Disconnected parallel PnP connection."
            }
            catch {
                Write-LogParallel "Error disconnecting parallel PnP connection: $_" "ERROR"
            }
        } -ThrottleLimit $DegreeOfParallelism -ArgumentList $_, $LocalBackupRoot, $DegreeOfParallelism, $SiteUrl, $ClientId, $TenantId, $ClientSecret, $LogFilePath, $Script:LogLock
    }

    # Disconnect from SharePoint Online
    try {
        Disconnect-PnPOnline -ErrorAction Stop
        Write-Log "Disconnected from SharePoint Online."
    }
    catch {
        Write-Log "Error disconnecting from SharePoint Online: $_" "ERROR"
    }

    Write-Log "Backup script completed."


# $SubFolderRelativeUrl = "/sites/Mercury/mercury_agreement"
# $SiteUrl = "https://morganhuntuk.sharepoint.com/sites/Mercury"
# $LocalBackupRoot = "e:\SharepointExportedDocs\test"
# $ClientId = "7ab7395c-1a70-4a6c-b623-2b2f28790fa5"
# $TenantId = "bd9ff961-8307-4028-b500-505726ed7ac4"
# $ClientSecret = "your-client-secret-here"  # Replace with your actual Client Secret
# $LogFilePath = "e:\SharepointExportedDocs\test\Logs\output.log"

# .\BackupAll-SharePoint.ps1 `
#     -SiteUrl $SiteUrl `
#     -LocalBackupRoot $LocalBackupRoot `
#     -ClientId $ClientId `
#     -TenantId $TenantId `
#     -ClientSecret $ClientSecret `
#     -SubFolderServerRelativeUrl $SubFolderRelativeUrl `
#     -LogFilePath $LogFilePath `
#     -DegreeOfParallelism 5
