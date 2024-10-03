# Command to run: .\MoveFoldersWithoutUnderscore.ps1 -ParentFolder "E:\Data\ParentFolder" -DestinationFolder "E:\Data\NoUnderscore" -ThrottleLimit 30


[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, HelpMessage = "Path to the parent folder.")]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$ParentFolder,

    [Parameter(Mandatory = $true, HelpMessage = "Path to the destination folder.")]
    [string]$DestinationFolder,

    [Parameter(Mandatory = $false, HelpMessage = "Number of parallel threads.")]
    [ValidateRange(1, 100)]
    [int]$ThrottleLimit = 20
)

# Start of the script
Write-Host "Starting the move process." -ForegroundColor Cyan
Write-Host "Parent Folder: $ParentFolder" -ForegroundColor Cyan
Write-Host "Destination Folder: $DestinationFolder" -ForegroundColor Cyan
Write-Host "Throttle Limit: $ThrottleLimit" -ForegroundColor Cyan

# Ensure DestinationFolder exists
if (-not (Test-Path -Path $DestinationFolder -PathType Container)) {
    try {
        New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
        Write-Host "Created destination folder: $DestinationFolder" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create destination folder: $DestinationFolder. Error: $_"
        exit 1
    }
}

# Define a log file path
$LogFile = Join-Path -Path $DestinationFolder -ChildPath "MoveFoldersLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

# Start logging
"Move Folders Log - $(Get-Date)" | Out-File -FilePath $LogFile -Encoding UTF8
"Parent Folder: $ParentFolder" | Out-File -FilePath $LogFile -Append
"Destination Folder: $DestinationFolder" | Out-File -FilePath $LogFile -Append
"Throttle Limit: $ThrottleLimit" | Out-File -FilePath $LogFile -Append
"------------------------------------------------------------" | Out-File -FilePath $LogFile -Append

try {
    # Retrieve all immediate subfolders
    Get-ChildItem -Path $ParentFolder -Directory -ErrorAction Stop | ForEach-Object -Parallel {
        param(
            $folder,
            $DestinationFolder,
            $LogFile
        )

        # Check if the folder name does NOT contain an underscore
        if ($folder.Name -notmatch '_') {
            try {
                # Define the target path
                $TargetPath = Join-Path -Path $using:DestinationFolder -ChildPath $folder.Name

                # Check if a folder with the same name already exists in the destination
                if (Test-Path -Path $TargetPath) {
                    # If exists, append a timestamp to the folder name to avoid conflicts
                    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
                    $TargetPath = Join-Path -Path $using:DestinationFolder -ChildPath ("{0}_{1}" -f $folder.Name, $Timestamp)
                }

                # Move the folder
                Move-Item -Path $folder.FullName -Destination $TargetPath -Force -ErrorAction Stop

                # Log the successful move
                $LogEntry = "Moved: $($folder.FullName) -> $TargetPath"
                Write-Output $LogEntry | Out-File -FilePath $using:LogFile -Append
            }
            catch {
                # Log the error
                $ErrorEntry = "Failed to move folder: $($folder.FullName). Error: $_"
                Write-Warning $ErrorEntry
                Write-Output $ErrorEntry | Out-File -FilePath $using:LogFile -Append
            }
        }
    } -ThrottleLimit $ThrottleLimit -ArgumentList $_, $DestinationFolder, $LogFile
}
catch {
    Write-Error "An error occurred during processing: $_"
}

Write-Host "Move process completed. Log file located at: $LogFile" -ForegroundColor Cyan
