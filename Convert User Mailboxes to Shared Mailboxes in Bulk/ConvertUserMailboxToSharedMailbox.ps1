
<#
=============================================================================================
Name:           Convert user mailbox to shared mailboxes in bulk
Description:    This script converts user mailboxes to shared mailboxes via import CSV
Version:        1.0
Website:        m365scripts.com
Blog:           https://m365scripts.com/exchange-online/microsoft-365-convert-user-mailbox-to-shared-mailbox-using-powershell
============================================================================================
#>


#Input File path declaration
$CSVPath=<FilePath>

#Check for EXO module inatallation
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
   Write-Host EXO module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 
Write-Host Connecting to Exchange Online...
Connect-ExchangeOnline  -ShowBanner:$false
Import-CSV $CSVPath | foreach {
 $UPN=$_.UPN
 Write-Progress -Activity "Converting $UPN to shared mailbox… "
 Set-Mailbox –Identity $UPN -Type Shared 
 If($?) 
 { 
  Write-Host $UPN Successfully converted to shared mailbox -ForegroundColor cyan 
 } 
 Else 
 { 
  Write-Host $UPN - Error occurred –ForegroundColor Red 
 } 
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 