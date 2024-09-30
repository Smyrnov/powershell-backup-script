# SharePoint Site URL
$siteUrl = "https://yourtenant.sharepoint.com/sites/yoursite"

# Document Library Name
$libraryName = "Documents"  # Change to your document library name

# Date Field to Filter By (e.g., "Created" or "Modified")
$dateField = "Created"

# Fields to Include in the View
# Add or remove fields as necessary. Ensure field internal names are used.
$viewFields = @("Title", "Name", $dateField)

# View Row Limit
$rowLimit = 100  # Adjust as needed

# View Type: "Standard" | "Calendar" | "Datasheet" etc.
$viewType = "Standard"

# Set as Default View
$setAsDefault = $false

# ------------------------------
# Script Execution
# ------------------------------

# Import PnP PowerShell Module
Import-Module PnP.PowerShell

# Connect to SharePoint Online
Write-Host "Connecting to SharePoint site: $siteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive

# Define Start and End Dates
$startDate = Get-Date "2020-01-01"
$endDate = Get-Date "2024-09-30"

# Initialize Current Date for Iteration
$currentDate = $startDate

# Loop Through Each Month
while ($currentDate -le $endDate) {
    # Extract Year and Month Information
    $year = $currentDate.Year
    $monthNumber = $currentDate.Month
    $monthName = $currentDate.ToString("MMMM")
    
    # Define View Name (e.g., "January 2020")
    $viewName = "$monthName $year"

    # Define Start and End of the Month
    $monthStart = $currentDate.ToString("yyyy-MM-dd")
    $nextMonth = $currentDate.AddMonths(1)
    $monthEnd = $nextMonth.ToString("yyyy-MM-dd")

    # Construct CAML Query for Filtering
    $camlQuery = @"
<View>
  <Query>
    <Where>
      <And>
        <Geq>
          <FieldRef Name='$dateField' />
          <Value IncludeTimeValue='FALSE' Type='DateTime'>$monthStart</Value>
        </Geq>
        <Lt>
          <FieldRef Name='$dateField' />
          <Value IncludeTimeValue='FALSE' Type='DateTime'>$monthEnd</Value>
        </Lt>
      </And>
    </Where>
  </Query>
</View>
"@

    # Check if the View Already Exists
    $existingView = Get-PnPView -List $libraryName -Identity $viewName -ErrorAction SilentlyContinue

    if ($null -eq $existingView) {
        try {
            # Create the View
            Add-PnPView -List $libraryName `
                       -Title $viewName `
                       -Fields $viewFields `
                       -RowLimit $rowLimit `
                       -Paged:$true `
                       -ViewType $viewType `
                       -Query $camlQuery `
                       -SetAsDefault:$setAsDefault

            Write-Host "Created view: $viewName" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to create view: $viewName. Error: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "View already exists: $viewName" -ForegroundColor Yellow
    }

    # Move to the Next Month
    $currentDate = $nextMonth
}

# Disconnect from SharePoint Online
Disconnect-PnPOnline
Write-Host "Script execution completed." -ForegroundColor Green
