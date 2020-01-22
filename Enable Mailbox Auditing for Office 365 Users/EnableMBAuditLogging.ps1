<# Purpose      : Enable mailbox audit logging for all Office 365 mailboxes
   Last updated : Jan 20, 2020
   Website      : https://O365reports.com
   For execution steps and usecases: https://o365reports.com/2020/01/21/enable-mailbox-auditing-in-office-365-powershell
#>

#Accept input paramenters 
param( 
[Parameter(Mandatory = $false)]
[string]$UserName, 
[string]$Password, 
[ValidateSet('ApplyRecord','Copy','Create','FolderBind','HardDelete','MessageBind','Move','MoveToDeletedItem','RecordDelete','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules','MailboxLogin')]
[string[]]$Operations=('ApplyRecord','Copy','Create','FolderBind','HardDelete','MessageBind','Move','MoveToDeletedItem','RecordDelete','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules','MailboxLogin'),
[switch]$MFA 
) 
 #Remove existing sessions
 Get-PSSession | Remove-PSSession 

 #Authentication using MFA 
 if($MFA.IsPresent) 
 { 
  $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1) 
  If ($MFAExchangeModule -eq $null) 
  { 
   Write-Host  `nPlease install Exchange Online MFA Module.  -ForegroundColor yellow 
   Write-Host You can install module using below blog: https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/ `n `nOR you can install module directly by entering "Y"`n 
   $Confirm= Read-Host Are you sure you want to install module directly? [Y] Yes [N] No 
   if($Confirm -match "[y]") 
   { 
    Write-Host Yes 
    Start-Process "iexplore.exe" "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application" 
   } 
   else 
   { 
    Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/' 
    Exit 
   } 
   $Confirmation= Read-Host Have you installed Exchange Online MFA Module? [Y] Yes [N] No 
    
   if($Confirmation -match "[y]") 
   { 
    $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1) 
    If ($MFAExchangeModule -eq $null) 
    { 
     Write-Host Exchange Online MFA module is not available -ForegroundColor red 
     Exit 
    } 
   } 
   else 
   {  
    Write-Host Exchange Online PowerShell Module is required 
    Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/' 
    Exit 
   }     
  } 
  
  #Importing Exchange MFA Module 
  Write-Host `nConnecting to Exchange Online...
  . "$MFAExchangeModule" 
  Connect-EXOPSSession -WarningAction SilentlyContinue 
  
 } 

 #Authentication using non-MFA 
 else 
 { 
  #Storing credential in script for scheduling purpose/ Passing credential as parameter 
  if(($UserName -ne "") -and ($Password -ne "")) 
  { 
   $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force 
   $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword 
  } 
  else 
  { 
   $Credential=Get-Credential -Credential $null 
  }  
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection 
  Write-Host `nConnecting to Exchange Online...
  Import-PSSession $Session -AllowClobber -DisableNameChecking| Out-Null 
 } 
 $MBCount=0
 $AuditAdmin="ApplyRecord","Copy","Create","FolderBind","HardDelete","MessageBind","Move","MoveToDeletedItem","RecordDelete","SendAs","SendOnBehalf","SoftDelete","Update","UpdateCalendarDelegation","UpdateFolderPermissions","UpdateInboxRules"
 $AuditDelegate ="ApplyRecord","Create","FolderBind","HardDelete","Move","MoveToDeletedItems","RecordDelete","SendAs","SendOnBehalf","SoftDelete","Update","UpdateFolderPermissions","UpdateInboxRules"
 $AuditOwner="ApplyRecord","Create","HardDelete","MailboxLogin","Move","MoveToDeletedItems","RecordDelete","SoftDelete","Update","UpdateCalendarDelegation","UpdateFolderPermissions","UpdateInboxRules"
 
if($Operations.Length -eq 17)
{
 $RequiredOperations=$Operations
 Get-Mailbox -ResultSize Unlimited | Select PrimarySmtpAddress,DisplayName | ForEach { 
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
 
 Get-Mailbox -ResultSize Unlimited | Select PrimarySmtpAddress,DisplayName | ForEach { 
  $DisplayName=$_.Displayname
  Write-Progress -Activity "`n     Processed mailbox count: $MBCount "`n"  Currently Processing: $DisplayName"
  $MBCount++
  Set-Mailbox -Identity $_.PrimarySmtpAddress -AuditEnabled $true -AuditAdmin $EnableAuditAdmin -AuditDelegate $EnableAuditDelegate -AuditOwner $EnableAuditowner
 }
}
Write-Host `nMailbox Audit logging enabled for $MBCount mailboxes -ForegroundColor Green
Write-Host "Mailbox Audit Logging enabled following operation(s):" $RequiredOperations
