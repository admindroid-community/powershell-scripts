<#
=============================================================================================

Name         : How to Find Which Retention Policies Are Applied to SharePoint Sites   
Version      : 1.0
website      : m365scripts.com

-----------------
Script Highlights
-----------------
1. Identifies which retention policies are applied to SharePoint Online sites.
2. Supports checking a single site, multiple sites via CSV input, or all sites in the tenant.
3. Filters and reports only enabled retention policies for accurate results.
4. Automatically verifies and installs required PowerShell modules (Exchange Online & SharePoint Online) upon your confirmation.
5. Supports Certificate-based Authentication (CBA) for unattended or secure automation.
6. Exports results to timestamped CSV files for easy tracking and archival.
7. Scheduler-friendly and suitable for automated compliance reporting. 

For detailed Script execution: https://m365scripts.com/sharepoint-online/how-to-find-which-retention-policies-are-applied-to-sharepoint-sites
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$Organization,
    [string]$ClientId,
    [string]$TenantId,
    [string]$CertificateThumbprint,
    [string]$HostName,
    [string]$SitesCSV,
    [string]$SiteURL
   
)

Function Connect_Exo
{
 #Check for EXO module installation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
   Import-Module ExchangeOnlineManagement
  } 
  else 
  { 
   Write-Host EXO module is required to connect Purview portal.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 
 Write-Host Connecting to Purview Compliance...
 
 
 if($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
   Connect-IPPSSession -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
 }
 else
 {
  Connect-IPPSSession -ShowBanner:$false
 }
}

Function Connect_SPO
{
 #Check for SPO module installation
 $SPOService = (Get-Module Microsoft.Online.SharePoint.PowerShell -ListAvailable).Name
 if ($SPOService -eq $null) 
 {
  Write-host "Important: SharePoint Online Management Shell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."  
  $confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
  if ($confirm -match "[Y]") 
  { 
   Write-host `n"Installing SharePoint Online Management Shell Module"
   Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Allowclobber -Repository PSGallery -Force -Scope CurrentUser
   Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
  }
  else 
  { 
   Write-host "Exiting. `nNote: SharePoint Online Management Shell module must be available in your system to run the script."  
   Exit 
  }
 }
 #Connecting to SharePoint Online PowerShell
 if($HostName -eq "")
 {
  Write-Host SharePoint organization name is required.`nEg: Contoso for admin@Contoso.Onmicrosoft.com -ForegroundColor Yellow
  $HostName= Read-Host "Please enter SharePoint organization name"  
 }
 $ConnectionUrl = "https://$HostName-admin.sharepoint.com/"  
 Write-Host `n"Connecting SharePoint Online Management Shell..."`n

 if($TenantId -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
  Connect-SPOService -Url $ConnectionUrl -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
}
 else
 {   
  Connect-SPOService -Url $ConnectionUrl | Out-Null
 }
}

Function Check_AppliedPolicies
{
 $Global:Count++
 $AppliedPolicies = @()
 Write-Progress -Activity "`n  Processed site count: $global:Count .."`n" Currently processing: $siteUrl"
 foreach ($policy in $policies) {
  $PolicyName=$Policy.Name
  $locations=@()
  $Exceptions=@()
  $locations  = $policy.SharePointLocation.Name
  $exceptions = $policy.SharePointLocationException.Name
  $isMatched = $false
  #Case 1: Policy applied to All SharePoint sites
  if($locations -contains "All") 
  {
   if($exceptions -notcontains $siteUrl) 
   {
    $AppliedPolicies += $PolicyName
   }
  }

  #Case 2: Policy applied to specific sites
  elseif ($locations -contains $siteUrl) 
  {
   $AppliedPolicies += $PolicyName
  }   
 }
 $AppliedRPCount= ( $AppliedPolicies| Measure-Object).Count
 if ($AppliedRPCount -eq 0) 
 {
  $AppliedPolicies = "No Policy Applied"
 }
    
 $Result= [PSCustomObject]@{
        'Site Url'  = $siteUrl
        'Applied Retention Policies' = ($AppliedPolicies -join ",")
        'Applied Policies Count'=$AppliedRPCount
    }
 $Result | Export-Csv $ExportCSV -NoTypeInformation -Append
}

#Connect to Required PowerShell services
Connect_Exo
if($SitesCSV -eq "" -and $SiteURL -eq "")
{
 Connect_SPO
}

$Location=Get-Location
$ExportCSV = "$Location\Sites_and_their_Retention_Policies_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$Policies = Get-RetentionCompliancePolicy -DistributionDetail | Where-Object {$_.SharePointLocation -ne $null -and $_.Enabled -eq $true}
$global:Count=0

Write-Host "Retrieving retention policy details for sites..."
#Process single site
if($SiteURL -ne "")
{
 Check_AppliedPolicies
}

#Process input CSV
elseif($SitesCSV -ne "")
{ 
 $Sites = Import-Csv $SitesCSV
 foreach ($Site in $sites) 
 { 
  $SiteURL = $Site.SiteUrl
  Check_AppliedPolicies
 }
}

#Process all sites
else
{
 (Get-SPOSite).url | foreach {
  $SiteURL=$_
  Check_AppliedPolicies
 }
}

#Open output file after execution

if((Test-Path -Path $ExportCSV) -eq "True") 
{
  Write-Host ""
  Write-Host " The Output file availble in:" -NoNewline -ForegroundColor Yellow
  Write-Host $ExportCSV
    Write-Host `nThe output file contains $global:Count site details.
  Write-Host `n~~ The Script is prepared by AdminDroid Community ~~`n -ForegroundColor Green 
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 3500+ Microsoft 365 reports & 450+ management actions. ~~" -ForegroundColor Green `n`n 
   $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$ExportCSV"   
  } 
 }
 


#Disconnect PowerShell sessions
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
if($SitesCSV -eq "" -and $SiteURL -eq "")
{
 Disconnect-SPOService
}
