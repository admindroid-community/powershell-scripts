<#
=============================================================================================
Name: Get SharePoint Files & Folders Created By External Users Using PowerShell
Version: 1.0
Website: o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the PnP module (if not installed already) upon your confirmation. 
2. Retrieves all files and folders created by external users for all sites. 
3. Gets files and folders created by external users on a specific site. 
4. Finds files and folders created by a specific external user. 
5. Allows to filter the data to display either files or folders created by external users. 
6. The script can be executed with an MFA-enabled account too. 
7. The script supports Certificate-based authentication (CBA) too. 
8. Exports the report results to a CSV file.

For detailed script execution: https://o365reports.com/2024/06/11/get-sharepoint-files-folders-created-by-external-users-using-powershell/

=============================================================================================
#>

param
( 
   [Parameter(Mandatory = $false)]
   [Switch] $FoldersOnly,
   [Switch] $FilesOnly,
   [string] $CreatedBy ,
   [String] $UserName,
   [String] $Password,
   [String] $ClientId,
   [String] $CertificateThumbprint,
   [String] $TenantName,  #(Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )
   [String] $SiteAddress,    #(Enter the specific site URL that you want to retrieve the data from.)
   [String] $SitesCsv 
)

#Check for SharePoint PnPPowerShellOnline module availability
$PnPOnline = (Get-Module PnP.PowerShell -ListAvailable).Name
if($PnPOnline -eq $null)
{ 
  Write-Host "Important: SharePoint PnP PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No  
  if($Confirm -match "[yY]")
  { 
    Write-Host "Installing SharePoint PnP PowerShell module..." -ForegroundColor Magenta
    Install-Module PnP.Powershell -Repository PsGallery -Force -AllowClobber -Scope CurrentUser
    Import-Module PnP.Powershell -Force
    #Register a new Azure AD Application and Grant Access to the tenant
    Register-PnPManagementShellAccess
  } 
  else
  { 
    Write-Host "Exiting. `nNote: SharePoint PnP PowerShell module must be available in your system to run the script" 
    Exit 
  }  
}

#Connecting to  SharePoint PnPPowerShellOnline module.......
Write-Host "Connecting to SharePoint PnPPowerShellOnline module..." -ForegroundColor Cyan
function Connect_Sharepoint
{
    param
    (
        [Parameter(Mandatory = $true)]
        [String] $Url
    )
    try
    {
        if(($UserName -ne "") -and ($Password -ne "") -and ($TenantName -ne ""))
        {
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
            Connect-PnPOnline -Url $Url -Credential $Credential
        }
        elseif($TenantName -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
        {
            Connect-PnPOnline -Url $Url -ClientId $ClientId -Thumbprint $CertificateThumbprint  -Tenant "$TenantName.onmicrosoft.com" 
        }
        else
        {
            Connect-PnPOnline -Url $Url -Interactive    
        }
    }
    catch
    {
        Write-Host "Error occured $($Url) : $_.Exception.Message"   -Foreground Red;
    }
}
if($TenantName -eq "")
{
    $TenantName = Read-Host "Enter your Tenant Name to Connect SharePoint Online  (Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )  "
}

$AdminUrl = "https://$TenantName.sharepoint.com/"
connect_sharepoint -Url $AdminUrl
$OutputCSV = "./SPO - Files & Folders Created By External Users " + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"

#Collecting the data and exporting it to a CSV file
$global:Count = 0
function Export_Data
{
    param
    (
        [Object] $ListItem,
        [Object] $ExternalUserIds,
        [String] $SiteUrl,
        [String] $SiteTitle
    ) 
    $AuthorFieldValue = $ListItem.FieldValues["Author"]
    $AuthorId = $AuthorFieldValue.LookupId
    $AuthorName = $AuthorFieldValue.LookupValue
    #Checking the resource created by an external user
    if(($ExternalUserIds | where{($_.Id -eq $AuthorId )}).count -eq 1)
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
        $ExportResult | Select-Object 'Site Name','Site Url','File/Folder Name','Created By','Resource Type','Created On','Relative URL' | Export-Csv -path $OutputCSV -Append -NoTypeInformation
    } 
}
#Collecting items created by external users
function Get_ExternalUserItems
{
    param
    (
        [String] $ObjectType,
        [String] $SiteUrl
    ) 
    try 
    {
        $Web = Get-PnPWeb | Select Title
        if($CreatedBy -eq "")
        {
            #Geting external Users present in site
            $ExternalUserIds = Get-PnPUser | where{($_.IsShareByEmailGuestUser -eq "True" -or $_.IsHiddenInUI -eq "True" ) } | Select Id 
        }
        else
        {
            
            $ExternalUserIds = Get-PnPUser | where{($_.IsShareByEmailGuestUser -eq "True" -or $_.IsHiddenInUI -eq "True" ) -and ($_.Email -eq $CreatedBy -or ($_.LoginName -split {$_ -eq "|"})[2] -eq $CreatedBy) } | Select Id
        }
        if(($ExternalUserIds).count -gt 0)
        {
            Get-PnPList | Where-Object {$_.Hidden -eq $false -and $_.BaseType -eq "DocumentLibrary"} | ForEach-Object{
                if($ObjectType -eq "All"){
                    # Retrieves list items
                    Get-PnPListItem -List $_.Title -PageSize 2000 | ForEach-Object{
                        Export_Data -ListItem $_ -ExternalUserIds $ExternalUserIds -SiteUrl $SiteUrl -SiteTitle $Web.Title
                    }
                }
                else
                {
                    # Retrieves list items for a specific object type.
                    Get-PnPListItem -List $_.Title -PageSize 2000 | where { $_.FileSystemObjectType -eq $ObjectType} |ForEach-Object{
                        Export_Data -ListItem $_ -ExternalUserIds $ExternalUserIds -SiteUrl $SiteUrl -SiteTitle $Web.Title
                    }
                }   
                
            }
        }
    }
    catch
    {
        Write-Host "Error occured $($SiteUrl) : $_"   -Foreground Red;
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

#To Retrive Data From All Sites Present In The Tenant
if($SiteAddress -ne "")
{
    Connect_Sharepoint -Url $SiteAddress 
    Get_ExternalUserItems -Objecttype $ObjectType -SiteUrl $SiteAddress
}
elseif($SitesCsv -ne "")
{
    try
    {
        Import-Csv -path $SitesCsv | ForEach-Object{
            Write-Progress -activity "Processing $($_.SitesUrl)" 
            Connect_Sharepoint -Url $_.SitesUrl 
            Get_ExternalUserItems -Objecttype $ObjectType -SiteUrl $_.SitesUrl
        }
    }
    catch
    {
        Write-Host "Error occured : $_"   -Foreground Red;
    }
}  
#To retrive the data for site presesent in our admin center
else
{
    Get-PnPTenantSite | Select Url,Title | ForEach-Object{
        Write-Progress -activity "Processing $($_.Url)" 
        Connect_Sharepoint -Url $_.Url 
        Get_ExternalUserItems -Objecttype $ObjectType -SiteUrl $_.Url
    }
}
#Open output file after execution
if($Count -gt 0)
{
    if((Test-Path -Path $OutputCSV) -eq "True")
    { 
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
    Write-Host "No records found"
}
#Disconnect the sharePoint PnPOnline module
Disconnect-PnPOnline
