# SharePoint Site URL
$siteUrl = ""
$ClientId = ""
$TenantId = ""

# Document Library Name
$libraryName = ""

# Local Download Base Path
$downloadBasePath = ""

# Date Range for Backup (Modify as needed)
$overallStartDate = Get-Date ""
$overallEndDate = Get-Date ""  # Exclusive

# CAML Query Parameters
$dateField = ""

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

# Connect to SharePoint Online
try {
    Write-Log "Connecting to SharePoint Online..."
    Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Tenant $TenantId -Interactive
    Write-Log "Successfully connected to SharePoint Online."
}
catch {
    Write-Log "Failed to connect to SharePoint Online: $_" -Level "ERROR" -Color Red
    exit 1
}

# Function to process items
function ProcessItems {
    param (
        [DateTime]$startDate,
        [DateTime]$endDate
    )

    # Log the current time step
    Write-Log "Starting processing for time step: $($startDate.ToString('yyyy-MM-dd HH:mm')) to $($endDate.ToString('yyyy-MM-dd HH:mm'))" -Color Cyan

    # Construct CAML Query with ViewFields
    $camlQuery = @"
<View Scope='RecursiveAll'>
    <ViewFields>
        <FieldRef Name='FileRef'/>
        <FieldRef Name='FileLeafRef'/>
        <FieldRef Name='FileDirRef'/>
        <FieldRef Name='FSObjType'/>
        <FieldRef Name='Created'/>
        <FieldRef Name='Modified'/>
    </ViewFields>
    <Query>
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
    </Query>
    <RowLimit Paged="TRUE">5000</RowLimit>
</View>
"@

    # Retrieve Items
    try {
        Write-Log "Executing CAML query for the specified date range..." -Color Cyan
        $items = Get-PnPListItem -List $libraryName -Query $camlQuery -PageSize 5000
        $itemCount = $items.Count
        Write-Log "CAML query returned $itemCount items." -Color Green
    }
    catch {
        Write-Log "Error retrieving items: $_" -Level "ERROR" -Color Red
        return
    }

    if (!$items -or $items.Count -eq 0) {
        Write-Log "No items found in the specified date range." -Color Yellow
        return
    }

    # Process files
    foreach ($item in $items) {
        $fileRef = $item.FieldValues["FileRef"]
        $fileLeafRef = $item.FieldValues["FileLeafRef"]
        $fileDirRef = $item.FieldValues["FileDirRef"]
        $fileSystemObjectType = $item.FieldValues["FSObjType"]

        # Skip folders
        if ($fileSystemObjectType -eq "1") { continue }

        # Construct the local file path
        $relativePath = $fileDirRef.Substring($fileDirRef.IndexOf($libraryName) + $libraryName.Length).TrimStart('/','\')
        $localFolderPath = Join-Path $downloadBasePath $relativePath
        $localFilePath = Join-Path $localFolderPath $fileLeafRef

        # Ensure the local folder exists
        if (!(Test-Path -Path $localFolderPath)) {
            try {
                New-Item -ItemType Directory -Path $localFolderPath -Force | Out-Null
                Write-Log "Created directory: $localFolderPath" -NoConsole
            }
            catch {
                Write-Log "❌ Failed to create directory: $localFolderPath. Error: $_" -Level "ERROR" -Color Red -NoConsole
                continue
            }
        }

        # Check if the file already exists with the same creation date
        $downloadFile = $true

        if (Test-Path -Path $localFilePath) {
            $localFile = Get-Item $localFilePath
            $createdDate = $item.FieldValues["Created"]

            if ($localFile.CreationTime -eq $createdDate) {
                # Log to file only
                Write-Log "File already exists and is up-to-date: $localFilePath" -Level "INFO" -NoConsole
                $downloadFile = $false
            }
        }

        if ($downloadFile) {
            try {
                Get-PnPFile -Url $fileRef -Path $localFolderPath -FileName $fileLeafRef -AsFile -Force
                # Log to file only
                Write-Log "Downloaded file: $localFilePath" -Level "INFO" -NoConsole

                # Set file dates
                $createdDate = $item.FieldValues["Created"]
                $modifiedDate = $item.FieldValues["Modified"]
                $localFile = Get-Item $localFilePath
                $localFile.CreationTime = $createdDate
                $localFile.LastWriteTime = $modifiedDate
            }
            catch {
                # Log to file only
                Write-Log "Error downloading file ${fileLeafRef}: $_" -Level "ERROR" -Color Red -NoConsole
            }
        }
    }

    Write-Log "Completed processing for time step: $($startDate.ToString('yyyy-MM-dd HH:mm')) to $($endDate.ToString('yyyy-MM-dd HH:mm'))" -Color Green
}

# Process the entire date range
$currentStartDate = $overallStartDate
$timeStepHours = 4

while ($currentStartDate -lt $overallEndDate) {
    $currentEndDate = $currentStartDate.AddHours($timeStepHours)
    if ($currentEndDate -gt $overallEndDate) {
        $currentEndDate = $overallEndDate
    }

    # Log the current time step to both log file and terminal
    Write-Log "Processing date range: $($currentStartDate.ToString('yyyy-MM-dd HH:mm')) to $($currentEndDate.ToString('yyyy-MM-dd HH:mm'))" -Color Cyan

    # Call the ProcessItems function
    ProcessItems -startDate $currentStartDate -endDate $currentEndDate

    # Introduce a delay to prevent throttling
    Start-Sleep -Seconds 5

    # Move to the next date range
    $currentStartDate = $currentEndDate
}

# Disconnect from SharePoint Online
Disconnect-PnPOnline
Write-Log "All date ranges processed successfully." -Color Green
