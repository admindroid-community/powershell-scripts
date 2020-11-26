#Accept input paramenters
param(
[switch]$FullAccess,
[switch]$SendAs,
[switch]$SendOnBehalf,
[string]$MBNamesFile,
[string]$UserName,
[string]$Password,
[switch]$MFA
)


function Print_Output
{
 #Print Output
 if($Print -eq 1)
 {
  $Result = @{'Display Name'=$_.Displayname;'User PrinciPal Name'=$upn;'Primary SMTP Address'=$PrimarySMTPAddress;'Access Type'=$AccessType;'User With Access'=$userwithAccess;'Email Aliases'=$EmailAlias}
  $Results = New-Object PSObject -Property $Result
  $Results |select-object 'Display Name','User PrinciPal Name','Primary SMTP Address','Access Type','User With Access','Email Aliases' | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

#Getting Mailbox permission
function Get_MBPermission
{
 $upn=$_.UserPrincipalName
 $DisplayName=$_.Displayname
 $MBType=$_.RecipientTypeDetails
 $PrimarySMTPAddress=$_.PrimarySMTPAddress
 $EmailAddresses=$_.EmailAddresses
 $EmailAlias=""
 foreach($EmailAddress in $EmailAddresses)
 {
  if($EmailAddress -clike "smtp:*")
  {
   if($EmailAlias -ne "")
   {
    $EmailAlias=$EmailAlias+","
   }
   $EmailAlias=$EmailAlias+($EmailAddress -Split ":" | Select-Object -Last 1 )
  }
 }
 $Print=0
 Write-Progress -Activity "`n     Processed mailbox count: $SharedMBCount "`n"  Currently Processing: $DisplayName"

 #Getting delegated Fullaccess permission for mailbox
 if(($FilterPresent -ne $true) -or ($FullAccess.IsPresent))
 {
  $FullAccessPermissions=(Get-MailboxPermission -Identity $upn | where { ($_.AccessRights -contains "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -match "NT AUTHORITY" -or $_.User -match "S-1-5-21") }).User
  if([string]$FullAccessPermissions -ne "")
  {
   $Print=1
   $UserWithAccess=""
   $AccessType="FullAccess"
   foreach($FullAccessPermission in $FullAccessPermissions)
   {
    if($UserWithAccess -ne "")
    {
     $UserWithAccess=$UserWithAccess+","
    }
    $UserWithAccess=$UserWithAccess+$FullAccessPermission
   }
   Print_Output
  }
 }

 #Getting delegated SendAs permission for mailbox
 if(($FilterPresent -ne $true) -or ($SendAs.IsPresent))
 {
  $SendAsPermissions=(Get-RecipientPermission -Identity $upn | where{ -not (($_.Trustee -match "NT AUTHORITY") -or ($_.Trustee -match "S-1-5-21"))}).Trustee
  if([string]$SendAsPermissions -ne "")
  {
   $Print=1
   $UserWithAccess=""
   $AccessType="SendAs"
   foreach($SendAsPermission in $SendAsPermissions)
   {
    if($UserWithAccess -ne "")
    {
     $UserWithAccess=$UserWithAccess+","
    }
    $UserWithAccess=$UserWithAccess+$SendAsPermission
   }
   Print_Output
  }
 }

 #Getting delegated SendOnBehalf permission for mailbox
 if(($FilterPresent -ne $true) -or ($SendOnBehalf.IsPresent))
 {
  $SendOnBehalfPermissions=$_.GrantSendOnBehalfTo
  if([string]$SendOnBehalfPermissions -ne "")
  {
   $Print=1
   $UserWithAccess=""
   $AccessType="SendOnBehalf"
   foreach($SendOnBehalfPermissionDN in $SendOnBehalfPermissions)
   {
    if($UserWithAccess -ne "")
    {
     $UserWithAccess=$UserWithAccess+","
    }
    #$SendOnBehalfPermission=(Get-Mailbox -Identity $SendOnBehalfPermissionDN).UserPrincipalName
    $UserWithAccess=$UserWithAccess+$SendOnBehalfPermissionDN
   }
   Print_Output
  }
 }
}

function main{
 #Connect AzureAD and Exchange Online from PowerShell
 Get-PSSession | Remove-PSSession

 #Authentication using MFA
 if($MFA.IsPresent)
 {
  $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
  If ($MFAExchangeModule -eq $null)
  {
   Write-Host  `nPlease install Exchange Online MFA Module.  -ForegroundColor yellow
   Write-Host You can install module using below blog :https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/ `n `nOR you can install module directly by entering "Y"`n
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
  . "$MFAExchangeModule"
  Connect-EXOPSSession -WarningAction SilentlyContinue
  Write-Host `nReport generation in progress...
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
  Import-PSSession $Session -CommandName Get-Mailbox,Get-MailboxPermission,Get-RecipientPermission -FormatTypeName * -AllowClobber | Out-Null
 }

 #Set output file
 $ExportCSV=".\SharedMBPermissionReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
 $Result=""
 $Results=@()
 $SharedMBCount=0
 $RolesAssigned=""

 #Check for AccessType filter
 if(($FullAccess.IsPresent) -or ($SendAs.IsPresent) -or ($SendOnBehalf.IsPresent))
 {
  $FilterPresent=$true
 }

 #Check for input file
 if ($MBNamesFile -ne "")
 {
  #We have an input file, read it into memory
  $MBs=@()
  $MBs=Import-Csv -Header "DisplayName" $MBNamesFile
  foreach($item in $MBs)
  {
   Get-Mailbox -Identity $item.displayname | Foreach{
   if($_.RecipientTypeDetails -ne 'SharedMailbox')
   {
     Write-Host $_.UserPrincipalName is not a shared mailbox -ForegroundColor Red
     continue
   }
   $SharedMBCount++
   Get_MBPermission
   }
  }
 }
 #Getting all Shared mailbox
 else
 {
  Get-mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | foreach{ 
   $SharedMBCount++
   Get_MBPermission}
 }


 #Open output file after execution
 Write-Host `nScript executed successfully
 if((Test-Path -Path $ExportCSV) -eq "True")
 {
  Write-Host "Detailed report available in: $ExportCSV"  -ForegroundColor Green
  $Prompt = New-Object -ComObject wscript.shell
  $UserInput = $Prompt.popup("Do you want to open output file?",`
 0,"Open Output File",4)
  If ($UserInput -eq 6)
  {
   Invoke-Item "$ExportCSV"
  }
 }
 Else
 {
  Write-Host No shared mailbox found that matches your criteria.
 }
#Clean up session
Get-PSSession | Remove-PSSession
}
 . main
