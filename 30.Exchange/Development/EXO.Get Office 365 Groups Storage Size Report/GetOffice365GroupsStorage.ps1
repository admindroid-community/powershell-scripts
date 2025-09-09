<#
=============================================================================================
Name:           Get storage used by Office 365 groups
Description:    This script find Office 365 groups' size and exports the report to CSV file
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script uses modern authentication to connect to Exchange Online.  
2. The script can be executed with MFA enabled account too. 
3. Automatically install the EXO V2 and SharePoint PnP PowerShell module (if not installed already) upon your confirmation.  
4. Credentials are passed as parameters (Scheduler-friendly), so worry not! i.e., credentials can be passed as parameters rather than being saved inside the script.  
5. Exports the report result to a CSV file. 
6. Lists the details of the storage used in each Office 365 group. 

For detailed script execution: https://o365reports.com/2022/05/18/get-the-storage-used-by-office-365-groups-using-powershell/
============================================================================================
#>
#PARAMETERS
param
( 
   [Parameter(Mandatory = $false)]
   [Switch] $NoMFA,
   [String] $UserName = $null, 
   [String] $Password = $null,
   [String] $TenantName = $null #(Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )
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
 

#Check for ExchangeOnline module availability
$Exchange = (Get-Module ExchangeOnlineManagement -ListAvailable).Name
if ($Exchange -eq $null)
{
  Write-Host "Important: Exchange Online PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."  
  $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
  if ($Confirm -match "[yY]") 
  { 
    Write-Host "Installing ExchangeOnlineManagement module" -ForegroundColor Magenta
    Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
    Import-Module ExchangeOnlineManagement -Force
  }
  else
  { 
    Write-Host "Exiting. `nNote: ExchangeOnline PowerShell module must be available in your system to run the script." 
    Exit 
  }
}


#Connecting to ExchangeOnline And SharePoint PnPPowerShellOnline module.......
Write-Host "Connecting to ExchangeOnline and SharePoint PnPPowerShellOnline module..." -ForegroundColor Cyan
#Authentication using non-MFA
if($NoMFA.IsPresent)
{
 #Storing credential in script for scheduling purpose/ Passing credential as parameter
 if(($UserName -ne "") -and ($Password -ne ""))
 { 
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
 }
 else
 {
  $Credential= Get-Credential -Credential $null
 }
 if($TenantName -eq "")
 {
  $TenantName = Read-Host "Enter your Tenant Name to complete the proccess (Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )  "
 }
  
  Connect-PnPOnline -Url https://$TenantName.sharepoint.com/ -Credentials $Credential
  Connect-ExchangeOnline -Credential $Credential
}
#Authentication using MFA
else
{
  $TenantName = Read-Host "Enter your Tenant Name to complete the proccess (Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )  " 
  Connect-PnPOnline -Url https://$TenantName.sharepoint.com/ -Interactive   
  Connect-ExchangeOnline 
}




#Get storage used by office 365 groups...
Write-Host "Getting office 365 groups storage..."`n
$OutputCsv=".\Office365GroupsStorageSizeReport_$((Get-Date -format MMM-dd` hh-mm` tt).ToString()).csv"
#Getting all sites which have an underlying Microsoft 365 group
$GroupSites = Get-PnPTenantSite -GroupIdDefined $true | Select-Object StorageUsageCurrent, StorageQuota, Url  
$GroupCount = 0

Get-UnifiedGroup -ResultSize unlimited | ForEach-Object {
 
 $GroupName = $_.DisplayName
 Write-Progress -Activity "Processed Group Count : $GroupCount" "Currently Processing Group : $GroupName"
 $SharePointSiteUrl = $_.SharePointSiteUrl
 if($SharePointSiteUrl -ne $null)
 {   
   $GroupSite = $GroupSites | Where-Object { $_.Url -eq $SharePointSiteUrl } 
   $GroupStorage = @{'Group Name' = $GroupName; 'Group Email' = $_.PrimarySmtpAddress; 'Group Privacy' = $_.AccessType;'Storage Used (GB)' = [math]::round($GroupSite.StorageUsageCurrent/1024,4); 'Storage Limit (GB)' = $GroupSite.StorageQuota/1024;  'Created On' = $_.WhenCreated}  
 }
 else
 {
   $GroupStorage = @{'Group Name' = $GroupName; 'Group Email' = $_.PrimarySmtpAddress; 'Group Privacy' = $_.AccessType;'Storage Used (GB)' = "Group not used yet"; 'Storage Limit (GB)' = "Group not used yet";  'Created On' = $_.WhenCreated} 
 }
 $ExportGroupStorage = New-Object PSObject -Property $GroupStorage
 $ExportGroupStorage | Select-Object 'Group Name','Group Email','Group Privacy','Storage Used (GB)','Storage Limit (GB)','Created On' | Export-Csv -path $OutputCsv -NoType -Append
 $GroupCount++
}

#Groupcount details
if($GroupCount -ne 0)
{
 Write-Host "$GroupCount Office 365 groups found in this organization."`n
}
else
{
 Write-Host "There is no office 365 group found in your organization"
}


#Open output file after execution 
if((Test-Path -Path $OutputCsv) -eq "True") 
{ 
  Write-Host " The office 365 groups storage report available in:" -NoNewline -ForegroundColor Yellow; Write-Host $OutputCsv 
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
  Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  $Prompt = New-Object -ComObject wscript.shell    
  $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)    
  If ($UserInput -eq 6)    
  {    
   Invoke-Item "$OutputCsv"    
  }  
} 

#Disconnect the SharePoint PnPPowerShellOnline module
Disconnect-PnPOnline

#Clean up session
Get-PSSession | Remove-PSSession