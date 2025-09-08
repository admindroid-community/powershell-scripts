<#
=============================================================================================
Name: Get SharePoint Files & Folders Created By External Users Using PowerShell
Version: 2.0
Website: o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the PnP module (if not installed already) upon your confirmation. 
2. Retrieves all files and folders created by external users for all sites. 
3. Gets files and folders created by external users on a specific site. 
4. Finds files and folders created by a specific external user. 
5. Allows to filter the data to display either files or folders created by external users. 
6. The script can be executed with an MFA-enabled account too. 
7. The script supports Certificate-based authentication (CBA) too. 
8. Automatic PnP App Registration handling with pre-configured Client ID
9. Exports the report results to a CSV file.

For detailed script execution: https://o365reports.com/2024/06/11/get-sharepoint-files-folders-created-by-external-users-using-powershell/

=============================================================================================
#>

param
( 
   [Parameter(Mandatory = $false)]
   [Switch] $FoldersOnly,
   [Switch] $FilesOnly,
   [string] $CreatedBy,
   [String] $UserName,
   [SecureString] $Password,
   [String] $ClientId,
   [String] $CertificateThumbprint,
   [String] $TenantName,  #(Example: If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name)
   [String] $SiteAddress, #(Enter the specific site URL that you want to retrieve the data from.)
   [String] $SitesCsv,
   [Switch] $Help
)

if ($Help) {
    Write-Host @"
SYNOPSIS
    Get SharePoint Files & Folders Created By External Users Using PowerShell

DESCRIPTION
    This script retrieves all files and folders created by external users in SharePoint Online.
    The script automatically handles PnP App Registration using a pre-configured Client ID.

PARAMETERS
    -FoldersOnly        : Show only folders created by external users
    -FilesOnly          : Show only files created by external users  
    -CreatedBy          : Filter by specific external user email
    -UserName           : Username for authentication
    -Password           : Password for authentication
    -ClientId           : Client ID for app authentication (default: afe1b358-534b-4c96-abb9-ecea5d5f2e5d)
    -CertificateThumbprint : Certificate thumbprint for authentication
    -TenantName         : Tenant name (e.g., 'contoso' for contoso.com)
    -SiteAddress        : Specific site URL to scan
    -SitesCsv          : CSV file with list of sites to scan
    -Help              : Show this help message

AUTHENTICATION
    The script uses a pre-registered PnP App with Client ID: afe1b358-534b-4c96-abb9-ecea5d5f2e5d
    If this app is not available, the script will attempt to register a new one automatically.

EXAMPLES
    # Show help
    .\ListSPOItemsCreatedbyExternalUsers.ps1 -Help

    # Scan all sites (requires admin permissions)
    .\ListSPOItemsCreatedbyExternalUsers.ps1

    # Scan specific site
    .\ListSPOItemsCreatedbyExternalUsers.ps1 -SiteAddress "https://tenant.sharepoint.com/sites/sitename"

    # Show only folders with custom Client ID
    .\ListSPOItemsCreatedbyExternalUsers.ps1 -FoldersOnly -ClientId "your-client-id" -SiteAddress "https://tenant.sharepoint.com/sites/sitename"

    # Use certificate authentication
    .\ListSPOItemsCreatedbyExternalUsers.ps1 -TenantName "yourtenant" -ClientId "your-client-id" -CertificateThumbprint "cert-thumbprint"
"@ -ForegroundColor Cyan
    exit 0
}

# Default PnP App Registration Client ID
$DefaultPnPClientId = "afe1b358-534b-4c96-abb9-ecea5d5f2e5d"

#Check for SharePoint PnPPowerShellOnline module availability
$PnPOnline = (Get-Module PnP.PowerShell -ListAvailable).Name
if($null -eq $PnPOnline)
{ 
  Write-Host "Important: SharePoint PnP PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
  $Confirm= Read-Host "Are you sure you want to install module? [Y] Yes [N] No: "
  if($Confirm -match "[yY]")
  { 
    Write-Host "Installing SharePoint PnP PowerShell module..." -ForegroundColor Magenta
    Install-Module PnP.Powershell -Repository PsGallery -Force -AllowClobber -Scope CurrentUser
    Import-Module PnP.Powershell -Force
  } 
  else
  { 
    Write-Host "Exiting. `nNote: SharePoint PnP PowerShell module must be available in your system to run the script" 
    Exit 
  }  
}

# Function to check and register PnP App if needed
function Initialize-PnPApp {
    param(
        [string]$TenantName
    )
    
    try {
        # If no ClientId provided in parameters, use the default one
        if ([string]::IsNullOrEmpty($script:ClientId)) {
            $script:ClientId = $DefaultPnPClientId
            Write-Host "Using default PnP Client ID: $DefaultPnPClientId" -ForegroundColor Yellow
        }
        
        # Test if the app registration works by attempting a connection
        $testUrl = "https://$TenantName.sharepoint.com/"
        Write-Host "Testing PnP App registration..." -ForegroundColor Cyan
        
        try {
            Connect-PnPOnline -Url $testUrl -ClientId $script:ClientId -Interactive -WarningAction SilentlyContinue
            Write-Host "✓ PnP App registration is working correctly!" -ForegroundColor Green
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            return $true
        }
        catch {
            Write-Host "⚠ PnP App registration test failed. Attempting to register..." -ForegroundColor Yellow
            
            # Register PnP Management Shell Access
            Write-Host "Registering PnP Management Shell Access..." -ForegroundColor Magenta
            $registration = Register-PnPManagementShellAccess
            
            if ($registration) {
                Write-Host "✓ PnP Management Shell registered successfully!" -ForegroundColor Green
                # Update ClientId if registration returned one
                if ($registration.ClientId) {
                    $script:ClientId = $registration.ClientId
                    Write-Host "New Client ID registered: $($script:ClientId)" -ForegroundColor Cyan
                }
                return $true
            } else {
                Write-Host "✗ Failed to register PnP Management Shell" -ForegroundColor Red
                return $false
            }
        }
    }
    catch {
        Write-Host "Error during PnP App initialization: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

#Connecting to SharePoint PnPPowerShellOnline module
Write-Host "Connecting to SharePoint PnPPowerShellOnline module..." -ForegroundColor Cyan

function Connect-SharePoint
{
    param
    (
        [Parameter(Mandatory = $true)]
        [String] $Url
    )
    try
    {
        if(($UserName -ne "") -and ($null -ne $Password) -and ($TenantName -ne ""))
        {
            $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$Password
            Connect-PnPOnline -Url $Url -Credential $Credential
        }
        elseif($TenantName -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
        {
            Connect-PnPOnline -Url $Url -ClientId $ClientId -Thumbprint $CertificateThumbprint -Tenant "$TenantName.onmicrosoft.com" 
        }
        elseif($ClientId -ne "")
        {
            # Use ClientId for interactive authentication
            Connect-PnPOnline -Url $Url -ClientId $ClientId -Interactive
        }
        else
        {
            # Fallback to default interactive connection
            Connect-PnPOnline -Url $Url -Interactive     
        }
        
        # Verify connection
        $web = Get-PnPWeb -ErrorAction Stop
        Write-Host "✓ Successfully connected to: $($web.Title)" -ForegroundColor Green
        return $true
    }
    catch
    {
        Write-Host "Error occurred connecting to $($Url) : $($_.Exception.Message)" -ForegroundColor Red;
        return $false
    }
}

if($TenantName -eq "")
{
    $TenantName = Read-Host "Enter your Tenant Name to Connect SharePoint Online (Example: If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name)"
}

# Initialize PnP App Registration
Write-Host "Initializing PnP App Registration..." -ForegroundColor Cyan
$pnpInitialized = Initialize-PnPApp -TenantName $TenantName
if (-not $pnpInitialized) {
    Write-Host "Failed to initialize PnP App Registration. Please check your permissions and try again." -ForegroundColor Red
    exit 1
}

# Check if we have a specific site address or if we need admin center access
$AdminUrl = "https://$TenantName.sharepoint.com/"

# Only connect to admin center if we're going to enumerate all sites
if($SiteAddress -eq "" -and $SitesCsv -eq "") {
    Write-Host "Connecting to SharePoint Admin Center..." -ForegroundColor Yellow
    $connected = Connect-SharePoint -Url $AdminUrl
    if (-not $connected) {
        Write-Host "Failed to connect to SharePoint Admin Center. You may not have admin permissions." -ForegroundColor Red
        Write-Host "Try running the script with -SiteAddress parameter to specify a specific site." -ForegroundColor Yellow
        exit 1
    }
}

$OutputCSV = "./SPO - Files & Folders Created By External Users " + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"

#Collecting the data and exporting it to a CSV file
$global:Count = 0
function Export-Data
{
    param
    (
        [Object] $ListItem,
        [Object] $ExternalUserIds,
        [String] $SiteUrl,
        [String] $SiteTitle
    ) 
    $AuthorFieldValue = $ListItem.FieldValues["Author"]
    if ($null -eq $AuthorFieldValue) {
        return
    }
    $AuthorId = $AuthorFieldValue.LookupId
    $AuthorName = $AuthorFieldValue.LookupValue
    #Checking the resource created by an external user
    if(($ExternalUserIds | Where-Object {($_.Id -eq $AuthorId )}).count -eq 1)
    {
        $global:Count++
        $ExportResult =@{
            'File/Folder Name'  = $ListItem.FieldValues.FileLeafRef;
            'Relative URL' = $AdminUrl + $ListItem.FieldValues.FileRef;
            'Created On' = if ($ListItem.FieldValues.Created) {$ListItem.FieldValues.Created} else { "-" } ;
            'Created By' =  $AuthorName;
            'Resource Type' = $ListItem.FileSystemObjectType;
            'Site Name' = if ($SiteTitle) {$SiteTitle} else { "-" };
            'Site Url' =  $SiteUrl   
        }
        $ExportResult = New-Object PSObject -Property $ExportResult
        #Export result to csv
        $ExportResult | Select-Object 'Site Name','Site Url','File/Folder Name','Created By','Resource Type','Created On','Relative URL' | Export-Csv -Path $OutputCSV -Append -NoTypeInformation
    } 
}

#Collecting items created by external users
function Get-ExternalUserItems
{
    param
    (
        [String] $ObjectType,
        [String] $SiteUrl
    ) 
    try 
    {
        $Web = Get-PnPWeb | Select-Object Title
        if($CreatedBy -eq "")
        {
            #Getting external Users present in site
            $ExternalUserIds = Get-PnPUser | Where-Object {($_.IsShareByEmailGuestUser -eq $true -or $_.IsHiddenInUI -eq $true) } | Select-Object Id 
        }
        else
        {
            $ExternalUserIds = Get-PnPUser | Where-Object {($_.IsShareByEmailGuestUser -eq $true -or $_.IsHiddenInUI -eq $true) -and ($_.Email -eq $CreatedBy -or ($_.LoginName -split '\|')[2] -eq $CreatedBy) } | Select-Object Id
        }
        if(($ExternalUserIds).count -gt 0)
        {
            Get-PnPList | Where-Object {$_.Hidden -eq $false -and $_.BaseType -eq "DocumentLibrary"} | ForEach-Object{
                if($ObjectType -eq "All"){
                    # Retrieves list items
                    Get-PnPListItem -List $_.Title -PageSize 2000 | ForEach-Object{
                        Export-Data -ListItem $_ -ExternalUserIds $ExternalUserIds -SiteUrl $SiteUrl -SiteTitle $Web.Title
                    }
                }
                else
                {
                    # Retrieves list items for a specific object type.
                    Get-PnPListItem -List $_.Title -PageSize 2000 | Where-Object { $_.FileSystemObjectType -eq $ObjectType} | ForEach-Object{
                        Export-Data -ListItem $_ -ExternalUserIds $ExternalUserIds -SiteUrl $SiteUrl -SiteTitle $Web.Title
                    }
                }   
            }
        }
        else {
            Write-Host "No external users found in site: $SiteUrl" -ForegroundColor Yellow
        }
    }
    catch
    {
        Write-Host "Error occurred processing $($SiteUrl) : $($_.Exception.Message)" -ForegroundColor Red;
    }
}

if($FoldersOnly.IsPresent)
{
    $ObjectType = "Folder"
}
elseif($FilesOnly.IsPresent)
{
    $ObjectType = "File"
}
else
{
  $ObjectType = "All"
}

#To retrieve Data From All Sites Present In The Tenant
if($SiteAddress -ne "")
{
    $connected = Connect-SharePoint -Url $SiteAddress 
    if ($connected) {
        Get-ExternalUserItems -ObjectType $ObjectType -SiteUrl $SiteAddress
    }
}
elseif($SitesCsv -ne "")
{
    try
    {
        Import-Csv -Path $SitesCsv | ForEach-Object{
            Write-Progress -Activity "Processing $($_.SitesUrl)" 
            $connected = Connect-SharePoint -Url $_.SitesUrl 
            if ($connected) {
                Get-ExternalUserItems -ObjectType $ObjectType -SiteUrl $_.SitesUrl
            }
        }
    }
    catch
    {
        Write-Host "Error occurred processing CSV : $($_.Exception.Message)" -ForegroundColor Red;
    }
}  
#To retrieve the data for sites present in our admin center
else
{
    try {
        Get-PnPTenantSite | Select-Object Url,Title | ForEach-Object{
            Write-Progress -Activity "Processing $($_.Url)" 
            $connected = Connect-SharePoint -Url $_.Url 
            if ($connected) {
                Get-ExternalUserItems -ObjectType $ObjectType -SiteUrl $_.Url
            }
        }
    }
    catch {
        Write-Host "Error retrieving tenant sites. Make sure you have admin permissions and are properly connected." -ForegroundColor Red
    }
}

#Open output file after execution
if($Count -gt 0)
{
    if((Test-Path -Path $OutputCSV) -eq "True")
    { 
        Write-Host "Found $Count items created by external users. Results saved to: $OutputCSV" -ForegroundColor Green
        $Prompt = New-Object -ComObject wscript.shell    
        $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File", 4)  
        If ($UserInput -eq 6)    
        {
            Invoke-Item $OutputCSV  
        }  
    }
}
else
{
    Write-Host "No records found" -ForegroundColor Yellow
}

#Disconnect the sharePoint PnPOnline module
try {
    Disconnect-PnPOnline
    Write-Host "Disconnected from SharePoint Online" -ForegroundColor Cyan
}
catch {
    # Ignore disconnect errors
}
