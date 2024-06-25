<#------------------------------------------------------------------------------------------------------------------------------------
Name: Export a List of all Document Library in SPO Using PowerShell
Version: 1.0
Website: o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the PnP module (if not installed already) upon your confirmation. 
2. Retrieves all the document libraries in SharePoint Online along with their details. 
3. Provides a list of all document libraries in a single site. 
4. Allows to get all document libraries and their details for multiple sites. 
5. The script can be executed with an MFA-enabled account too. 
6. The script supports Certificate-based authentication (CBA) too. 
7. Exports the report results to a CSV file. 
8. The script is scheduler friendly.

For detailed script execution: https://o365rpeorts.com/2024/06/25/list-all-document-library-in-spo-using-powershell/
-----------------------------------------------------------------------------------------------------------------------------------
#>
Param
( 
   [Parameter(Mandatory = $false)]
   [String] $UserName , 
   [String] $Password ,
   [String] $ClientId,
   [String] $CertificateThumbprint,
   [String] $TenantName, #(Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )
   [String] $SiteAddress,   #(Enter the specific site URL that you want to retrieve the data from.)
   [String] $SitesCsv 
)
$PnPOnline = (Get-Module PnP.PowerShell -ListAvailable).Name
if($PnPOnline -eq $null)
{
  Write-Host "Important: SharePoint PnP PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No  
  if($Confirm -match "[yY]")
  {
    Write-Host "Installing SharePoint PnP PowerShell module..." -ForegroundColor Magenta
    Install-Module PnP.Powershell -Repository PsGallery -Force -AllowClobber 
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
function Connect_SharePoint
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
    $TenantName = Read-Host "Enter your Tenant Name to Connect to SharePoint Online (Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )  "
}

$AdminUrl = "https://$TenantName.sharepoint.com"
connect_sharepoint -Url $AdminUrl
$Location=Get-Location
$OutputCSV = "$Location\SPO Document Library Report " + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"

#Converting to the nearest Unit of size
function Convert_ToNearestUnit {
    param (
        [long]$LibrarySizeInBytes
    )
    if ($LibrarySizeInBytes -eq 0) {
        return "0 Bytes"
    }
    $Units = ("Bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    $NearestIndex = [Math]::Min([Math]::Floor([Math]::Log($LibrarySizeInBytes, 1024)), $Units.Count - 1)
    $SizeInNearestUnit = [Math]::Round($LibrarySizeInBytes / [Math]::Pow(1024, $NearestIndex), 2)
    $Unit = $Units[$NearestIndex]
    return "$sizeInNearestUnit $Unit"
}
#Collecting Document Reports
function Get_Statistics
{
    param
    (
        [String] $SiteUrl,
        [String] $SiteTitle
    ) 
    Get-PnPList  | Where-Object {$_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false} | ForEach-Object{
        if($_.Title -ne "Form Templates" -and $_.Title -ne "Style Library")
        {
            $LibrarySize = Get-PnPFolderStorageMetric -List $_.Title | Select TotalSize,TotalFileCount
            $FolderCount = $_.ItemCount - $LibrarySize.TotalFileCount
            $FilesCount = $LibrarySize.TotalFileCount
            $LibrarySizeInBytes = $LibrarySize.TotalSize
            $LibrarySize = Convert_ToNearestUnit -LibrarySizeInBytes $LibrarySizeInBytes
            $ExportResult = @{
                "Document Library Name" = $_.Title;
                "Document Library Url" = $AdminUrl+$_.DefaultViewUrl;
                "Created On" = $_.Created;
                "Site Url" =  $SiteUrl;
                "Site Name" = if ($SiteTitle) {$SiteTitle} else { "-" };
                "Library Size(Bytes)" = $LibrarySizeInBytes;
                "Folders Count" = $FolderCount;
                "Files Count" = $FilesCount;
                "Library Size" = $LibrarySize;
            }
            $ExportResult = New-Object PSObject -Property $ExportResult
            #Exporting data in to Csv 
            $ExportResult | Select-Object "Site Name","Site Url","Document Library Name","Document Library Url","Created On","Library Size","Library Size(Bytes)","Folders Count","Files Count" | Export-Csv -path $OutputCSV -Append -NoTypeInformation   
       
        }
    } 
}

#Retriving the data for site presesent in the tenant
if($SiteAddress -ne "")
{
    Connect_SharePoint -Url $SiteAddress
    $Web = Get-PnPWeb | Select Title,Url
    Get_Statistics -SiteUrl $Web.Url -SiteTitle $Web.Title   
}
#Retriving data for specified sites present in the tenant
elseif($SitesCsv -ne "")
{
    try
    {
        Import-Csv -path $SitesCsv | ForEach-Object{
            Write-Progress -activity "Processing $($_.SitesUrl)" 
            Connect_Sharepoint -Url $_.SitesUrl 
            $Web = Get-PnPWeb | Select Url,Title 
            Get_Statistics -Objecttype $ObjectType -SiteUrl $Web.Url -SiteTitle $Web.Title
        }
    }
    catch
    {
        Write-Host "Error occured : $_"   -Foreground Red;
    }
}  
#Retriving data from all sites present in the tenant
else
{
    Get-PnPTenantSite | Select Url,Title | ForEach-Object{
        Write-Progress -activity "Processing $($_.Url)" 
        Connect_SharePoint -Url $_.Url
        Get_Statistics -SiteUrl $_.Url -SiteTitle $_.Title
    }
}

#Open output file after execution
if((Test-Path -Path $OutputCSV) -eq "True") 
{   
    Write-Host `n "The Output file availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCSV" `n 
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    $Prompt = New-Object -ComObject wscript.shell    
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File", 4)   
     
    If ($UserInput -eq 6)
    {    
        Invoke-Item $OutputCSV    
    }  
}

#Disconnect the sharePoint PnPOnline module
Disconnect-PnPOnline