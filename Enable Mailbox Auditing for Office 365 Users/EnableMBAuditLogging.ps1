<# Purpose      : Enable mailbox audit logging for all Office 365 mailboxes
   Last updated : Feb 22, 2025
   Website      : https://O365reports.com
 
Script Highlights:
~~~~~~~~~~~~~~~~~~

1.You can enable mailbox audit logging for all mailboxes using single cmdlet 
2.This script allows you to enable mailbox auditing for all mailbox actions or set of actions 
3.The script can be executed with MFA enabled account also. 
4.This script automatically installs Exchange Online PowerShell Module upon your confirmation  
   
   For execution steps and usecases: https://o365reports.com/2020/01/21/enable-mailbox-auditing-in-office-365-powershell
#>

#Accept input paramenters 
param( 
[Parameter(Mandatory = $false)]
[string]$UserName, 
[string]$Password, 
 [string]$Organization,
 [string]$ClientId,
 [string]$CertificateThumbprint,
[ValidateSet('ApplyRecord','Copy','Create','FolderBind','HardDelete','MailItemsAccessed','MessageBind','Move','MoveToDeletedItems','RecordDelete','SearchQueryInitiated','Send','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateComplianceTag','UpdateFolderPermissions','UpdateInboxRules','MailboxLogin')]
[string[]]$Operations=('ApplyRecord','Copy','Create','FolderBind','HardDelete','MailItemAccessed','MessageBind','Move','MoveToDeletedItems','RecordDelete','SearchQueryInitiated','Send','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateComplianceTag','UpdateFolderPermissions','UpdateInboxRules','MailboxLogin')
) 


Function Connect_Exo
{
 #Check for EXO module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell  module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
  } 
  else 
  { 
   Write-Host EXO module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 
 Write-Host Connecting to Exchange Online...
 #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
 }
 elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
   Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
 }
 else
 {
  Connect-ExchangeOnline -ShowBanner:$false
 }
}
Connect_Exo 
 $MBCount=0
 $AuditAdmin="ApplyRecord","Copy","Create","FolderBind","HardDelete","MailItemsAccessed","Move","MoveToDeletedItems","RecordDelete","Send","SendAs","SendOnBehalf","SoftDelete","Update","UpdateCalendarDelegation","UpdateComplianceTag","UpdateFolderPermissions","UpdateInboxRules"
 $AuditDelegate ="ApplyRecord","Create","FolderBind","HardDelete","MailItemsAccessed","Move","MoveToDeletedItems","RecordDelete","SendAs","SendOnBehalf","SoftDelete","Update","UpdateComplianceTag","UpdateFolderPermissions","UpdateInboxRules"
 $AuditOwner="ApplyRecord","Create","HardDelete","MailItemsAccessed","MailboxLogin","Move","MoveToDeletedItems","RecordDelete","SearchQueryInitiated","Send","SoftDelete","Update","UpdateCalendarDelegation","UpdateComplianceTag","UpdateFolderPermissions","UpdateInboxRules"
 
if($Operations.Length -eq 21)
{
 $RequiredOperations=$Operations
 Get-EXOMailbox -ResultSize Unlimited | Select PrimarySmtpAddress,DisplayName | ForEach { 
  $DisplayName=$_.Displayname
  Write-Progress -Activity "`n     Processed mailbox count: $MBCount "`n"  Currently Processing: $DisplayName"
  $MBCount++
  Set-Mailbox -Identity $_.PrimarySmtpAddress -AuditEnabled $true -AuditAdmin $AuditAdmin -AuditDelegate $AuditDelegate -AuditOwner $Auditowner
 }
}
else
{
 $RequiredOperations=$PSBoundParameters.Operations
 [System.Collections.ArrayList]$EnableAuditAdmin=@()
 [System.Collections.ArrayList]$EnableAuditDelegate=@()
 [System.Collections.ArrayList]$EnableAuditOwner=@()
 Foreach($Operation in $RequiredOperations)
 {
  if($AuditAdmin -match $Operation)
  {
   $EnableAuditAdmin += $Operation
  }
  if($AuditDelegate -match $Operation)
  {
   $EnableAuditDelegate += $Operation
  }
  if($AuditOwner -match $Operation)
  {
   $EnableAuditOwner += $Operation
  }
 }
 
 Get-EXOMailbox -ResultSize Unlimited | Select PrimarySmtpAddress,DisplayName | ForEach { 
  $DisplayName=$_.Displayname
  Write-Progress -Activity "`n     Processed mailbox count: $MBCount "`n"  Currently Processing: $DisplayName"
  $MBCount++
  Set-Mailbox -Identity $_.PrimarySmtpAddress -AuditEnabled $true -AuditAdmin $EnableAuditAdmin -AuditDelegate $EnableAuditDelegate -AuditOwner $EnableAuditowner
 }
}
Write-Host `nMailbox Audit logging enabled for $MBCount mailboxes -ForegroundColor Yellow `n
Write-Host "Mailbox Audit Logging enabled following operation(s):" $RequiredOperations
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
