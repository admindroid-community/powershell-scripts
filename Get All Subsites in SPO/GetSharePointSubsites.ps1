<#
=============================================================================================
Name:           Get All Subsites in SharePoint Online Using PowerShell 
Version:        2.0
website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Exports subsites of all sites in the SharePoint tenant. 
2. Exports subsites for a list of sites alone. 
3. Retrieves subsites for a specific site based on user input. 
4. Shows the current recycle bin status (Enabled/Disabled) for sites. 
5. Automatically installs the PnP PowerShell module (if not installed already) upon your confirmation. 
6. The script can be executed with an MFA enabled account too. 
7. Exports report results as a CSV file. 
8. The script is scheduler friendly. 
9. The script uses modern authentication to connect SharePoint Online. 
10. It can be executed with certificate-based authentication (CBA) too. 

For detailed Script execution: https://o365reports.com/2024/06/04/get-all-subsites-in-sharepoint-online-using-powershell/

~~~~~~~~~
Change Log:
~~~~~~~~~
  V1.0 (Jun 04, 2024) - File created
  V2.0 (Dec 29, 2025) - Handled ClientId requirement for SharePoint PnP PowerShell module and made minor usability changes

============================================================================================
#>
Param
( 
   [Parameter(Mandatory = $false)]
   [String] $UserName, 
   [String] $Password,
   [String] $ClientId,
   [String] $CertificateThumbprint,
   [String] $TenantName, #(Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )
   [String] $SiteAddress,  #(Enter the specific site URL that you want to retrieve the data from.)
   [String] $SitesCsv 
)

#Check for SharePoint PnPPowerShellOnline module availability
Function Installation-Module
{
 $Module = Get-InstalledModule -Name PnP.PowerShell -MinimumVersion 1.12.0 -ErrorAction SilentlyContinue
 If($Module -eq $null)
 {
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
   Write-Host "PnP PowerShell module is required to connect SharePoint Online.Please install module using 'Install-Module PnP.PowerShell' cmdlet."
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
 if(($AdminName -ne "") -and ($Password -ne "") -and ($ClientId -ne ""))
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
  Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
 }
}


$global:Count = 0
function Get_Subsites
{
    param
    (
        [String] $SiteUrl
    ) 
    try
    {
        $Web = Get-PnPWeb | Select Title
        # Getting the subsites of site
        Get-PnPSubWeb -Recurse -Includes Created,LastItemUserModifiedDate,Description,RecycleBinEnabled | ForEach-Object{
            $ExportResult = [PSCustomObject][ordered]@{
                "Site Collection Name" = $web.Title
                "Site Collection Url" = $SiteUrl
                "Site Name"  = $_.Title            
                "Site URL" = $_.Url
                "Site description" = if ($_.Description) { $_.Description } else { "-" }
                "Creation Date" = $_.Created
                "Last Modified Date" = $_.LastItemUserModifiedDate
                "Recycle Bin Enabled" = if ($_.RecycleBinEnabled) { $_.RecycleBinEnabled } else { "-" }
            }
            #Exporting data in to Csv file
            $ExportResult | Export-Csv -Path $OutputCSV -Append -NoTypeInformation
            $global:Count++
        }
    }
    catch 
    {
        Write-Host "Error occured $($SiteUrl): $($_.Exception.Message)" -Foreground Yellow
    }
}


Installation-Module
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$OutputCSV = "$(Get-Location)\SPO_Subsites_Report_$timestamp.csv"

if($TenantName -eq "" -and $SiteAddress -eq "" -and $SitesCsv -eq "")
{
 $TenantName = Read-Host "Enter your Tenant Name to connect SharePoint Online (Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )  "
}

if($ClientId -eq "")
{
 $ClientId= Read-Host "ClientId is required to connect PnP PowerShell. Enter ClientId"
}

#Retriving data from all sites present in the tenant
if($SiteAddress -ne "")
{
    try{
        Connection-Module -Url $SiteAddress
        Get_Subsites -SiteUrl $SiteAddress
    }
    catch{
        $_.Exception.Message
    }
}
elseif($SitesCsv -ne "")
{
    try
    {
        Import-Csv -path $SitesCsv | ForEach-Object{
            Write-Progress -activity "Processing $($_.SitesUrl)" 
            Connection-Module -Url $_.SitesUrl 
            Get_Subsites -SiteUrl $_.SitesUrl
        }
    }
    catch
    {
        $_.Exception.Message
    }
}
#Retriving the data for site presesent in the tenant
else
{
    try{
        Connection-Module -Url "https://$TenantName-admin.sharepoint.com/"
        Get-PnPTenantSite | Where -Property Template -NotIn ("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1") | ForEach-Object{
            Write-Progress -activity "Processing $($_.Url)"
            Connection-Module -Url $_.Url
            Get_Subsites -SiteUrl $_.Url 
        }
    }
    catch{
        $_.Exception.Message
    }
}
#Open output file after execution
if($Count -gt 0)
{
    if((Test-Path -Path $OutputCSV) -eq "True") 
    { 
        Write-Host `nThe output file contains $Global:Count sites
        Write-Host "`n The Output file availble in: " -NoNewline -ForegroundColor Yellow; Write-Host $OutputCSV;
        Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
        Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host "to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green `n`n
        $Prompt = New-Object -ComObject wscript.shell    
        $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File", 4)    
        if ($UserInput -eq 6)    
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
Disconnect-PnPOnline -WarningAction SilentlyContinue