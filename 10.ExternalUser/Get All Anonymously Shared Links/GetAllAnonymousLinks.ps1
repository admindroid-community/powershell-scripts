<#
=============================================================================================
Name:           Get All Anonymous Links in SharePoint Online Using PowerShell  
Version:        1.0
Website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~ 
1. Exports all anonymous links in your SharePoint Online environment. 
2. The script automatically verifies and installs the PnP PowerShell module (if not installed already) upon your confirmation. 
3. Identifies files and folders with only active anonymous links. 
4. Lists files and folders that have only expired anonymous links. 
5. Generates a report that retrieves never expiring anyone links.
6. Exports a report that retrieves anonymous links set with expiration.
7. Allows to export the list of files and folders with soon to expire anyone links (i.e., 30 days, 90 days, etc.)  
8. The script can be executed with an MFA-enabled account too. 
9. Exports report results to CSV file. 
10. The script uses modern authentication to connect SharePoint Online. 
11. It can be executed with certificate-based authentication (CBA) too. 
12. The script is scheduler friendly.

~~~~~~~~~~~~~~~~~~
Note:
~~~~~~~~~~~~~~~~~~
Make sure the app registration used for certificate based authentication is granted with Application permissions
for "Files.Read.All" and "Sites.Read.All" to ensure it can retrieve the required details. Else, you will get the
following error:"Get-PnPFileSharingLink: Either scp or roles claim need to be present in the token."


For detailed script execution: https://o365reports.com/2025/04/22/get-all-anonymous-links-in-sharepoint-online-using-powershell

============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$AdminName ,
    [string]$Password ,
    [String]$ClientId ,
    [String]$CertificateThumbprint,
    [string]$TenantName,
    [string]$ImportCsv,
    [Switch]$ActiveLinks,
    [Switch]$ExpiredLinks,
    [Switch]$LinksWithExpiration,
    [Switch]$NeverExpiresLinks,
    [int]$SoonToExpireInDays 
)

#Check for module availability
Function Installation-Module
{
 $Module = Get-InstalledModule -Name PnP.PowerShell -MinimumVersion 1.12.0 -ErrorAction SilentlyContinue
 If($Module -eq $null){
 Write-Host SharePoint PnP PowerShell Module is not available -ForegroundColor Yellow
 $Confirm = Read-Host Are you sure you want to install module? [Yy] Yes [Nn] No
 If($Confirm -match "[yY]") 
 { 
  Write-Host "Installing PnP PowerShell module..."
  Install-Module PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
  Import-Module -Name Pnp.Powershell        
 } 
 Else
 { 
  Write-Host PnP PowerShell module is required to connect SharePoint Online.Please install module using Install-Module PnP.PowerShell cmdlet. 
  Exit
 }
}
    Write-Host `nConnecting to SharePoint Online...
}

#SPO Site connection 
Function Connection-Module
{
 param
 (
  [Parameter(Mandatory = $true)]
  [String] $Url
 )
 if(($AdminName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $AdminName,$SecuredPassword
  Connect-PnPOnline -Url $Url -Credential $Credential -ClientId $ClientId
 }
 elseif($TenantName -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
  Connect-PnPOnline -Url $Url -ClientId $ClientId -Thumbprint $CertificateThumbprint  -Tenant "$TenantName.onmicrosoft.com" 
 }
 else
 {
  Connect-PnPOnline -Url $Url -ClientId $ClientId -Interactive
 }
}

Function Get-SharedLinks
{
 $ExcludedLists = @("Form Templates","Style Library","Site Assets","Site Pages", "Preservation Hold Library", "Pages", "Images",
                       "Site Collection Documents", "Site Collection Images")
 $DocumentLibraries = Get-PnPList | Where-Object {$_.Hidden -eq $False -and $_.Title -notin $ExcludedLists -and $_.BaseType -eq "DocumentLibrary"} # -ErrorAction SilentlyContinue 
 Foreach($List in $DocumentLibraries){
 $ListItems = Get-PnPListItem -List $List -PageSize 2000 
 ForEach ($Item in $ListItems) 
 {
  $FileName=$Item.FieldValues.FileLeafRef
  $ObjectType=$Item.FileSystemObjectType
  Write-Progress -Activity ("Site Name: $($Site.Title)") -Status ("Processing Item: "+ $FileName )
  $HasUniquePermissions = Get-PnPProperty -ClientObject $Item -Property HasUniqueRoleAssignments
  
  If ($HasUniquePermissions) 
  {    
   $FileUrl=$Item.FieldValues.FileRef
   if($ObjectType -eq "File")
   {
    $FileSharingLinks= Get-PnPFileSharingLink -Identity $FileUrl
   }
   elseif($ObjectType -eq "Folder")
   {
    $FileSharingLinks= Get-PnPFolderSharingLink -Folder $FileUrl
   }
   else
   {
    continue
   }
   foreach ($FileSharingLink in $FileSharingLinks)
   {
    $Link=$FileSharingLink.Link

    #Process Anonymous links
    if($Link.Scope -eq "Anonymous")
    {
     $Permission=$Link.Type
     $SharedLink=$Link.WebUrl
     $PasswordProtected=$FileSharingLink.HasPassword 
     $BlockDownload=$Link.PreventsDownload
     $RoleList = $FileSharingLink.Roles -join ","
     $ExpirationDate=$FileSharingLink.ExpirationDateTime
     $CurrentDateTime = (Get-Date).Date
     If($ExpirationDate -ne $null)
     {
      $ExpiryDate = ([DateTime]$ExpirationDate).ToLocalTime()
      $ExpiryDays= (New-TimeSpan -Start $CurrentDateTime -End $ExpiryDate).Days
      If($ExpiryDate -lt $CurrentDateTime)
      {
       $LinkStatus = "Expired"
       $ExpiryDateCalculation=$ExpiryDays * (-1)
       $FriendlyExpiryTime="Expired $ExpiryDateCalculation days ago"
      } 
      else
      {
       $LinkStatus="Active"
       $FriendlyExpiryTime="Expires in $ExpiryDays days"
      }
     }
     else
     {
      $LinkStatus="Active"
      $ExpiryDays="-"
      $ExpiryDate="-"
      $FriendlyExpiryTime="Never Expires"
     }

     #Filter for active anonymous links
     If(($ActiveLinks.IsPresent) -and ($LinkStatus -ne "Active"))
     {
      Continue
     }
     
     #Filter for expired anonymoous links
     elseif(($ExpiredLinks.IsPresent) -and ($LinkStatus -ne "Expired"))
     {
      Continue
     }
     
     #Filter for finding anonymous links with expiration
     elseif(($LinksWithExpiration.IsPresent) -and ($ExpirationDate -eq $null))
     {
      Continue
     } 
     
     #Filter for finding anonymous links with no expiry
     elseif(($NeverExpiresLinks.IsPresent) -and ($FriendlyExpiryTime -ne "Never Expires"))
     {
      Continue
     }
     
     #Filter for finding soon-to-expire anonymous links
     elseif(($SoonToExpireInDays -ne "") -and (($ExpirationDate -eq $null) -or ($SoonToExpireInDays -lt $ExpiryDays) -or ($ExpiryDays -lt 0)))  
     {
      Continue
     }

     $Results = [PSCustomObject]@{
                            "Site Name"             = $Site.Title
                            "Library"          = $List.Title
                            "Object Type" =$ObjectType
                            "File/Folder Name"             = $FileName
                            "File/Folder URL"         = $FileUrl
                            "Access Type"      = $Permission
                            "Roles" =$RoleList
                            "File Type"         = $Item.FieldValues.File_x0020_Type 
                            "Link Status"     = $LinkStatus
                            "Link Expiry Date"  = $ExpiryDate
                            "Days Since/To Expiry" =$ExpiryDays
                            "Friendly Expiry Time" =$FriendlyExpiryTime
                            "Password Protected" =$PasswordProtected
                            "Block Download"    =$BlockDownload                   
                            "Shared Link"       = $SharedLink  
      }
      $Results | Export-CSV  -path $ReportOutput -NoTypeInformation -Append  -Force
      $Global:ItemCount++
     }                    
    }
   }
  }
 }
}


$TimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ReportOutput = "$PSScriptRoot\AnonymousLink_Report_ $TimeStamp.csv"
$Global:ItemCount = 0
If($ClientId -eq "")
{
 $ClientId= Read-Host "ClientId is required to connect PnP PowerShell. Enter ClientId"
}
If($TenantName -eq "")
{
    $TenantName = Read-Host "Enter your tenant name (e.g., 'contoso' for 'contoso.onmicrosoft.com')"
}

#Check for CSV input
If($ImportCsv -ne "")
{
 $SiteCollections = Import-Csv -Path $ImportCsv
 Foreach($Site in $SiteCollections){
  $SiteUrl = $Site.SiteUrl
  Connection-Module -Url $SiteUrl
  try {
    $Site = Get-PnPWeb 
    Get-SharedLinks
  }
  catch{
    Write-Host "You don't have access to this site : $SiteUrl" -ForegroundColor Yellow
  } 
 }
 Disconnect-PnPOnline -WarningAction SilentlyContinue
}

#Process all sites
Else
{
 Connection-Module -Url "https://$TenantName-admin.sharepoint.com"
 $SiteCollections = Get-PnPTenantSite  | Where -Property Template -NotIn ("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")
 ForEach($Site in $SiteCollections)
 {
  Connection-Module -Url $Site.Url
    try { 
         Get-SharedLinks
    }
    catch {
         Write-Host "You don't have access to this site : $($Site.Url)" -ForegroundColor Yellow
    }
 }
 Disconnect-PnPOnline -WarningAction SilentlyContinue
}
 Write-Progress -Completed
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 
if((Test-Path -Path $ReportOutput) -eq "True") 
{
    Write-Host `nThe output file contains $Global:ItemCount anonymous links.
    Write-Host `n The Output file availble in:  -NoNewline -ForegroundColor Yellow
    Write-Host $ReportOutput
      $Prompt = New-Object -ComObject wscript.shell   
    $UserInput = $Prompt.popup("Do you want to open output file?",`   
    0,"Open Output File",4)   
    If ($UserInput -eq 6)   
    {   
        Invoke-Item "$ReportOutput"   
    } 
}
else{
    Write-Host "No Records Found"
}