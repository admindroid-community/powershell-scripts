<#
=============================================================================================
Name:           Get All Subsites in SharePoint Online Using PowerShell 
Version:        1.0
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

#Connecting to share point....
write-Host "Connecting SharePoint PnPPowerShellOnline module..."
function Connect_Sharepoint
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [String]$Url
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
    $TenantName = Read-Host "Enter your Tenant Name to connect SharePoint Online (Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )  "
}
$AdminUrl = "https://$TenantName.sharepoint.com/"
Connect_Sharepoint -Url $AdminUrl
$OutputCSV = ".\Subsites Report " + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"

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
            $global:Count++
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
        }
    }
    catch 
    {
        Write-Host "Error occured $($SiteUrl) : $_"   -Foreground Red;
    }
} 


#Retriving data from all sites present in the tenant
if($SiteAddress -ne "")
{
    Connect_SharePoint -Url $SiteAddress
    Get_Subsites -SiteUrl $SiteAddress
}
elseif($SitesCsv -ne "")
{
    try
    {
        Import-Csv -path $SitesCsv | ForEach-Object{
            Write-Progress -activity "Processing $($_.SitesUrl)" 
            Connect_Sharepoint -Url $_.SitesUrl 
            Get_Subsites -SiteUrl $_.SitesUrl
        }
    }
    catch
    {
        Write-Host "Error occured : $_"   -Foreground Red;
    }
}
#Retriving the data for site presesent in the tenant
else
{
    Get-PnPTenantSite | Select Url,Title | ForEach-Object{
        Write-Progress -activity "Processing $($_.Url)" 
        Connect_SharePoint -Url $_.Url
        Get_Subsites -SiteUrl $_.Url 
    }
}
#Open output file after execution
if($Count -gt 0)
{
    if((Test-Path -Path $OutputCSV) -eq "True") 
    { 
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
Disconnect-PnPOnline