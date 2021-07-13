Param
(
    [Parameter(Mandatory = $false)]
    [switch]$CheckStatus,
    [switch]$Enable,
    [switch]$Disable
)
#Check for EXO v2 module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell V2 module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
  } 
  else 
  { 
   Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 

 #Connect Exchange Online
 Write-Host Connecting to Exchange Online... -ForegroundColor Cyan
 Connect-ExchangeOnline
 
 #Check for Plus Addressing status
 if($CheckStatus.IsPresent)
 {
  $Status=Get-OrganizationConfig | select AllowPlusAddressInRecipients 
  if($Status.AllowPlusAddressInRecipients -eq $true)
  {
   Write-Host Currently, Plus Addressing is enabled in your organization.
  }
  else
  {
   Write-Host Currently,Plus Addressing is disabled in your organization.
  }
 }

 #Enable Plus Addressing
 if($Enable.IsPresent)
 {
  Set-OrganizationConfig –AllowPlusAddressInRecipients $True 
  if($?)
  {
   Write-Host Plus addressing enabled successfully -ForegroundColor Yellow
  }
 }

 #Disable Plus Addressing
 if($Disable.IsPresent)
 {
  Set-OrganizationConfig –AllowPlusAddressInRecipients $False
  if($?)
  {
   Write-Host Plus addressing disabled successfully -ForegroundColor Yellow
  }
 }