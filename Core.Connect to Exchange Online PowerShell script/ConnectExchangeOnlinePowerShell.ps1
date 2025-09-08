<#
=============================================================================================
Name:           Connect to Exchange Online PowerShell
Version:        2.0
Website:        o365reports.com

For detailed script execution:  https://o365reports.com/2019/08/22/connect-exchange-online-powershell/
============================================================================================
#>
#Due RPS and Basic Auth retirement in Exchange Online,  we are no longer able to use modules earlier than EXO V3.
#Check for EXO v3 module installation
$Module = (Get-Module ExchangeOnlineManagement -ListAvailable) | where {$_.Version.major -ge 3}
if($Module.count -eq 0)
{
 Write-Host Exchange Online PowerShell module is not available -ForegroundColor yellow
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
 if($Confirm -match "[yY]")
 {
 Write-host "Installing Exchange Online PowerShell module"
 Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
 Import-Module ExchangeOnlineManagement
 }
 else
 {
 Write-Host EXO module is required to connect Exchange Online. Please install module using Install-Module ExchangeOnlineManagement cmdlet.
 Exit
 }
}
 
Write-Host `nConnecting to Exchange Online...
Connect-ExchangeOnline

Write-Host Script executed successfully!
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n