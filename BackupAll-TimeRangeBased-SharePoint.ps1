# SharePoint Site URL
$siteUrl = ""

# Document Library Name
$libraryName = ""

# Local Download Base Path
$downloadBasePath = ""

# Date Range for Backup (Modify as needed)
$overallStartDate = Get-Date "2021-01-01 00:00"
$overallEndDate = Get-Date "2021-04-10 00:00"  # Exclusive

# CAML Query Parameters
$dateField = "Created"

# Log file base path
$logFileBasePath = ""

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
        [string]$FolderServerRelativeUrl,
        [string]$LocalFolderPath
    )

    # Retrieve folder properties from SharePoint
    try {
        $folder = Get-PnPFolder -Url $FolderServerRelativeUrl -Includes TimeCreated, TimeLastModified
    }
    catch {
        Write-Log "❌ Failed to get properties for folder: $FolderServerRelativeUrl. Error: $_" -Level "ERROR" -Color Red -NoConsole
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
            # Optionally, you can decide to update the folder's creation date or skip it
            # For this script, we'll update the creation date
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

# Define the ProcessTimeRange function
function ProcessTimeRange {
    param (
        [DateTime]$startDate,
        [DateTime]$endDate,
        [int]$timeStepMinutes
    )

    # Log the date range being processed
    Write-Log "Processing date range: $($startDate.ToString('yyyy-MM-dd HH:mm')) to $($endDate.ToString('yyyy-MM-dd HH:mm')) with time step $timeStepMinutes minutes" -Color Cyan

    # Construct Explicit CAML Query for the Backup
    $camlQuery = @"
<Where>
    <And>
        <Geq>
            <FieldRef Name='$dateField' />
            <Value IncludeTimeValue='TRUE' Type='DateTime'>$($startDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"))</Value>
        </Geq>
        <Lt>
            <FieldRef Name='$dateField' />
            <Value IncludeTimeValue='TRUE' Type='DateTime'>$($endDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"))</Value>
        </Lt>
    </And>
</Where>
"@

    # Wrap the CAML Query within <View Scope='RecursiveAll'> to include all folders and subfolders
    $fullCamlQuery = @"
<View Scope='RecursiveAll'>
    <Query>
        $camlQuery
    </Query>
    <RowLimit>5000</RowLimit>
</View>
"@

    Write-Log "🔍 Retrieving items based on the defined CAML query..." -Color Cyan

    # Retrieve Items Based on the CAML Query
    try {
        $items = Get-PnPListItem -List $libraryName -Query $fullCamlQuery -PageSize 1000
        Write-Log "✅ Retrieved $($items.Count) items matching the CAML query." -Color Green
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Log "❌ Failed to retrieve items. Error: $errorMessage" -Level "ERROR" -Color Red

        if ($errorMessage -like "*exceeds the list view threshold*") {
            if ($timeStepMinutes -gt 1) {
                # Split the range into smaller ranges with smaller time steps
                if ($timeStepMinutes -gt 60) {
                    $newTimeStep = 60  # Next smaller step is 60 minutes (1 hour)
                } elseif ($timeStepMinutes -gt 1) {
                    $newTimeStep = 1   # Next smaller step is 1 minute
                }

                Write-Log "⚠️ Splitting date range into smaller intervals of $newTimeStep minutes due to list view threshold." -Level "WARN" -Color Yellow

                $tempStartDate = $startDate
                while ($tempStartDate -lt $endDate) {
                    $tempEndDate = $tempStartDate.AddMinutes($newTimeStep)
                    if ($tempEndDate -gt $endDate) {
                        $tempEndDate = $endDate
                    }
                    # Recursively call ProcessTimeRange with the smaller range
                    ProcessTimeRange -startDate $tempStartDate -endDate $tempEndDate -timeStepMinutes $newTimeStep
                    $tempStartDate = $tempEndDate
                }
            }
            else {
                # Cannot split further, log and skip
                Write-Log "⚠️ Unable to process date range: $($startDate.ToString('yyyy-MM-dd HH:mm')) to $($endDate.ToString('yyyy-MM-dd HH:mm')) due to list view threshold limitations." -Level "WARN" -Color Yellow
            }
        }
        else {
            # Other error, log and exit
            Write-Log "❌ Unexpected error occurred: $errorMessage" -Level "ERROR" -Color Red
            Disconnect-PnPOnline
            exit 1
        }

        # Since we handled the error, return from the function
        return
    }

    if ($items.Count -eq 0) {
        Write-Log "No items found for date range: $($startDate.ToString('yyyy-MM-dd HH:mm')) to $($endDate.ToString('yyyy-MM-dd HH:mm'))" -Color Yellow
        return
    }

    Write-Log "⬇️ Downloading items..." -Color Cyan

    foreach ($item in $items) {
        $fileRef = $item.FieldValues["FileRef"]
        $fileLeafRef = $item.FieldValues["FileLeafRef"]
        $fileDirRef = $item.FieldValues["FileDirRef"]
        $fileSystemObjectType = $item.FieldValues["FSObjType"]

        # Calculate the relative path within the library
        $libraryPath = "/$libraryName/"
        $startIndex = $fileDirRef.ToLower().IndexOf($libraryPath.ToLower())
        if ($startIndex -ge 0) {
            $relativePath = $fileDirRef.Substring($startIndex + $libraryPath.Length).TrimStart('/')
        }
        else {
            $relativePath = $fileDirRef.TrimStart('/')
        }

        # Extract the top-level folder name
        $relativePathParts = $relativePath.Split([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
        if ($relativePathParts.Count -gt 0) {
            $topLevelFolderName = $relativePathParts[0]
        } else {
            $topLevelFolderName = ""
        }

        # Check if the top-level folder name contains an underscore
        if ($topLevelFolderName -notmatch "_") {
            # Skip processing the item
            Write-Log "⏭ Skipping item as top-level folder '$topLevelFolderName' does not contain an underscore." -Color Yellow -NoConsole
            continue
        }

        $localFolderPath = Join-Path $downloadBasePath $relativePath

        # Ensure all parent folders exist and set their dates
        $folderPaths = @()
        $currentRelativePath = ""
        foreach ($part in $relativePathParts) {
            if ([string]::IsNullOrEmpty($currentRelativePath)) {
                $currentRelativePath = $part
            } else {
                $currentRelativePath = Join-Path $currentRelativePath $part
            }
            $folderPaths += $currentRelativePath
        }

        $baseServerRelativePath = "/$libraryName"
        foreach ($folderPath in $folderPaths) {
            # Construct the server-relative URL for the folder
            $folderServerRelativeUrl = "$baseServerRelativePath/$($folderPath -replace '\\','/')"
            $currentLocalFolderPath = Join-Path $downloadBasePath $folderPath

            Ensure-LocalFolder -FolderServerRelativeUrl $folderServerRelativeUrl -LocalFolderPath $currentLocalFolderPath
        }

        if ($fileSystemObjectType -eq "1") {
            # It's a folder (already handled above)
            continue
        }
        else {
            # It's a file
            # Download the file if it doesn't already exist with the same creation date
            try {
                $localFilePath = Join-Path $localFolderPath $fileLeafRef

                $createdDate = $item.FieldValues["Created"]
                $modifiedDate = $item.FieldValues["Modified"]

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
                        # For this script, we'll skip downloading the file
                        $downloadFile = $false
                    }
                }

                if ($downloadFile) {
                    Get-PnPFile -Url $fileRef -Path $localFolderPath -FileName $fileLeafRef -AsFile -Force
                    Write-Log "✅ Downloaded file: $fileLeafRef to $localFolderPath" -Color Green -NoConsole

                    # Set the file's creation and modification dates
                    (Get-Item $localFilePath).CreationTime = $createdDate
                    (Get-Item $localFilePath).LastWriteTime = $modifiedDate
                }
            }
            catch {
                Write-Log "❌ Failed to download file: $fileLeafRef. Error: $_" -Level "ERROR" -Color Red -NoConsole
            }
        }
    }

    Write-Log "✅ Completed processing for date range: $($startDate.ToString('yyyy-MM-dd HH:mm')) to $($endDate.ToString('yyyy-MM-dd HH:mm'))" -Color Green
}

# Connect to SharePoint Online
Write-Log "Connecting to SharePoint site: $siteUrl" -Color Cyan

try {
    Write-Log "Connecting to SharePoint Online..."
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive
    Write-Log "Successfully connected to SharePoint Online."
}
catch {
    Write-Log "Failed to connect to SharePoint Online: $_" -Level "ERROR" -Color Red
    exit 1
}

# Start of the sliding window loop
$currentStartDate = $overallStartDate

while ($currentStartDate -lt $overallEndDate) {
    $backupStartDate = $currentStartDate
    $backupEndDate = $currentStartDate.AddHours(4)  # End date is exclusive

    # Call the function to process this time range with initial time step of 240 minutes (4 hours)
    ProcessTimeRange -startDate $backupStartDate -endDate $backupEndDate -timeStepMinutes 240

    # Move to the next date range
    $currentStartDate = $currentStartDate.AddHours(4)
}

# Disconnect from SharePoint Online
Disconnect-PnPOnline
Write-Log "✅ All date ranges processed successfully." -Color Green
