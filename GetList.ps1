# SharePoint Site URL
$siteUrl = 
$ClientId = 
$TenantId = 
# Document Library Name
$libraryName = 

# Date Range for Folder Retrieval (Modify as needed)
$overallStartDate = Get-Date "2021-01-01 00:00"
$overallEndDate = Get-Date "2024-11-01 00:00"  # Exclusive

# CAML Query Parameters
$dateField = "Created"

# Output file path
$outputFilePath = 

# Log file base path
$logFileBasePath = 

# Import PnP PowerShell Module
Import-Module PnP.PowerShell -ErrorAction Stop

# Ensure the output file directory exists
$outputDir = Split-Path -Parent $outputFilePath
if (!(Test-Path -Path $outputDir)) {
    try {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    catch {
        Write-Host "❌ Failed to create output directory: $outputDir. Error: $_" -ForegroundColor Red
        exit 1
    }
}

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

# Initialize output file
Clear-Content -Path $outputFilePath -ErrorAction SilentlyContinue

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

# Define the ProcessTimeRange function
function ProcessTimeRange {
    param (
        [DateTime]$startDate,
        [DateTime]$endDate,
        [int]$timeStepMinutes
    )

    # Log the date range being processed
    Write-Log "Processing date range: $($startDate.ToString('yyyy-MM-dd HH:mm')) to $($endDate.ToString('yyyy-MM-dd HH:mm')) with time step $timeStepMinutes minutes" -Color Cyan

    # Construct Explicit CAML Query for the Folders
    $camlQuery = @"
<Where>
    <And>
        <Eq>
            <FieldRef Name='FSObjType' />
            <Value Type='Integer'>1</Value>
        </Eq>
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

    Write-Log "🔍 Retrieving folders based on the defined CAML query..." -Color Cyan

    # Retrieve Items Based on the CAML Query
    try {
        $items = Get-PnPListItem -List $libraryName -Query $fullCamlQuery -PageSize 1000
        Write-Log "✅ Retrieved $($items.Count) folders matching the CAML query." -Color Green
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Log "❌ Failed to retrieve folders. Error: $errorMessage" -Level "ERROR" -Color Red

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
        Write-Log "No folders found for date range: $($startDate.ToString('yyyy-MM-dd HH:mm')) to $($endDate.ToString('yyyy-MM-dd HH:mm'))" -Color Yellow
        return
    }

    Write-Log "Processing folders..." -Color Cyan

    foreach ($item in $items) {
        $folderRef = $item.FieldValues["FileRef"]

        # Write the folder path to the output file
        try {
            Add-Content -Path $outputFilePath -Value $folderRef
            Write-Log "✅ Wrote folder path to output file: $folderRef" -Color Green -NoConsole
        }
        catch {
            Write-Log "❌ Failed to write folder path: $folderRef. Error: $_" -Level "ERROR" -Color Red -NoConsole
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
    $backupEndDate = $currentStartDate.AddHours(24)  # End date is exclusive

    ProcessTimeRange -startDate $backupStartDate -endDate $backupEndDate -timeStepMinutes 1440

    # Move to the next date range
    $currentStartDate = $currentStartDate.AddHours(24)
}

# Disconnect from SharePoint Online
Disconnect-PnPOnline
Write-Log "✅ All date ranges processed successfully." -Color Green
