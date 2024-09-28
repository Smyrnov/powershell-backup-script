# Information
**Folders:** Only folders with an underscore in their names are processed.

**Document Libraries:** Only libraries with underscores in their names are processed.

**Files:** All files within the processed folders and libraries are handled based on the LastModifiedDate parameter.

**Optional filtering:** ensures that only files created after **CreatedDate** or modified after **LastModifiedDate** are downloaded, thereby optimizing the backup process by avoiding unnecessary file downloads when running backup for the same SharePoint instance multiple times.

## Requirements
- Entra ID Application Registration
    To register application run: 
    ```
    Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnPAppName" -Tenant "" -Interactive -SharePointDelegatePermissions "AllSites.Read"
    ```
    **Note 1:** "AllSites.Read" may be replaced by "AllSites.FullControl" if more access is needed.
    **Note 2:** You must have the necessary permissions to create app registrations in your Entra ID.

    After registering the application, you might need to grant admin consent for the permissions.
    Roles required for this are: Azure AD Global Administrator or Application Developer, or Cloud Application Administrator role.
    An administrator can grant consent by navigating to the **Azure Portal > Azure Active Directory > App registrations > Your application > API permissions > Grant admin consent** 
    Click on **"Grant admin consent for [Your Tenant Name]"**.
- Administrative Privileges - should be executed in PowerShell with "Run as Administrator" option
- PowerShell 7.0 or later
- .NET Runtime
- PnP.PowerShell Module. 
    + To install: ``Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber ``
    + To update: ``Update-Module PnP.PowerShell``
    + To verify: ``Get-Module PnP.PowerShell -ListAvailable | Select-Object Name, Version``

## Usage Instructions
1. Save the Script

2. Prepare MandatoryParameters

- SiteUrl: The URL of your SharePoint site (e.g., https://mysite.sharepoint.com/sites/sitename).
- LocalBackupRoot: The local directory where you want to store the backups (e.g., C:\SharePointBackup\sitename).
- ClientId: The application (client) ID of your registered Entra ID application.
- TenantId: Your tenant domain or tenant ID.

Optional Parameters:

- SubFolderServerRelativeUrl: The server-relative URL of the subfolder you want to back up (e.g., /sites/sitename/foldername).
- LastModifiedDate: A datetime value to determine if existing files should be overwritten based on modification date (e.g., 2024-09-25T00:00:00).
- CreatedDate: A datetime value to determine if files should be downloaded based on creation date (e.g., 2024-09-25T00:00:00).
- LogFilePath: Path to a custom log file. If not provided, a timestamped log file is created in Logs under LocalBackupRoot.

3. Run the Script
Open PowerShell 7 or later and execute the script with the required parameters.

#### Example 1: Backing Up a Specific Subfolder with LastModifiedDate

Define parameters
```
$SiteUrl = "https://mysite.sharepoint.com/sites/sitename"
$SubFolderServerRelativeUrl = "/sites/sitename/foldername"
$LocalBackupRoot = "C:\SharePointBackup\foldername"
$ClientId = "your-application-client-id"
$TenantId = "yourtenant.onmicrosoft.com"
# Optional:
$LastModifiedDate = Get-Date "2024-09-25T00:00:00"
$CreatedDate = Get-Date "2024-09-20T00:00:00"
$LogFilePath = "C:\SharePointBackup\Logs\foldername_BackupLog.txt"
```
Run the script
```
.\Backup-SharePoint.ps1 `
    -SiteUrl $SiteUrl `
    -SubFolderServerRelativeUrl $SubFolderServerRelativeUrl `
    -LocalBackupRoot $LocalBackupRoot `
    -ClientId $ClientId `
    -TenantId $TenantId `
    -LastModifiedDate $LastModifiedDate `
    -CreatedDate $CreatedDate `
    -LogFilePath $LogFilePath
```

#### Example 2: Backing Up the Entire Site Without Any Filtering

Define parameters
```
$SiteUrl = "https://mysite.sharepoint.com/sites/foldername"
$LocalBackupRoot = "C:\SharePointBackup\FullSiteBackup"
$ClientId = "your-application-client-id"
$TenantId = "yourtenant.onmicrosoft.com"
```
Run the script
```
.\Backup-SharePoint.ps1 `
    -SiteUrl $SiteUrl `
    -LocalBackupRoot $LocalBackupRoot `
    -ClientId $ClientId `
    -TenantId $TenantId
```

## Additional Resources


+ [PnP.PowerShell GitHub Repository](https://github.com/pnp/powershell)
+ [PnP.PowerShell Cmdlet Reference](https://pnp.github.io/powershell/cmdlets/)
+ [Azure AD App Registration Documentation](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app?tabs=certificate)
+ [SharePoint API Permissions Reference](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs)
+ [Microsoft Graph Permissions Reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
