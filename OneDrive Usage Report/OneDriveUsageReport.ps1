<#
=============================================================================================
Name:     Get OneDrive Usage Report in Microsoft 365      
Version:  1.0
Website:  o365reports.com   

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script allows you to export OneDrive usage report for all OneDrive sites.
2. The script uses SharePoint Online Management Shell module and installs it if not already present (with user confirmation).
3. Exports the report result to CSV.

For detailed script execution:https://o365reports.com/2025/09/16/get-onedrive-usage-report-in-microsoft-365/

============================================================================================
#>


param (
    [string] $UserName,
    [string] $Password,
    [string] $HostName
       
)

#Checks SharePointOnline module availability and connects the module
Function ConnectSPOService
{
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
 if (($UserName -ne "") -and ($Password -ne "") ) 
 {   
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force   
  $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword   
  Connect-SPOService -Credential $Credential -Url $ConnectionUrl | Out-Null
 }   
 else 
 {   
  Connect-SPOService -Url $ConnectionUrl | Out-Null
 }
 
}
ConnectSPOService
$Location=Get-Location
$ExportCSV="$Location\OneDriveStorageReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$Count=0
Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'" | foreach {
 $Count++
 $UserName=$_.Title
 $UPN=$_.Owner
 $Url=$_.Url
 Write-Progress -Activity "`n     Processed OneDrive site count: $Count "`n"  Currently processing site: $url"
 $StorageSize=$_.StorageUsageCurrent
 $LastContentModifiedDate=$_.LastContentModifiedDate
 $StorageQuota=$_.StorageQuota
 $StorageQuotainGB=[math]::round($StorageQuota/1024, 2)
 $StorageSizeinGB=[math]::round($StorageSize/1024, 2)
 $Status=$_.Status
 $ArchiveStatus=$_.ArchiveStatus
 $Result=@{ 'Owner UPN'=$UPN;'OneDrive Url'=$url;'Storage Used (GB)'=$StorageSizeinGB;'Storage Quota (GB)'=$StorageQuotainGB;'Storage Used (MB)'=$StorageSize;'Storage Quota (MB)'=$StorageQuota;'Status'=$Status;'Archive Status'=$ArchiveStatus;'Last Content Modified Date'=$LastContentModifiedDate}
 $Results= New-Object PSObject -Property $Result  
 $Results | Select-Object 'Owner UPN','OneDrive Url','Storage Used (GB)','Storage Quota (GB)','Storage Used (MB)','Storage Quota (MB)','Status','Archive Status','Last Content Modified Date'| Export-Csv -Path $ExportCSV -Notype -Append }


Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green `n`n

 

 if((Test-Path -Path $ExportCSV) -eq "True") 
 {
  Write-Host `The exported report contains $Count OneDrive sites.
  Write-Host `nOneDrive sites report available in: -NoNewline -Foregroundcolor Yellow; Write-Host $ExportCSV
   $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$ExportCSV"   
  } 
 }
 else
 {
  Write-Host No OneDrive accounts found.
 }