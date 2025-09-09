<#
=============================================================================================
Name:           Get Shared Mailbox Permission Report 
Version:        2.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~~

1.The script display only “Explicitly assigned permissions” to mailboxes which means it will ignore “SELF” permission that each user on his mailbox and inherited permission. 
2.Exports output to CSV file. 
3.The script can be executed with MFA enabled account also. 
4.You can choose to either “export permissions of all mailboxes” or pass an input file to get permissions of specific mailboxes alone. 
5.Allows you to filter output using your desired permissions like Send-as, Send-on-behalf or Full access. 
6.This script is scheduler friendly. I.e., credentials can be passed as a parameter instead of saving inside the script 

For detailed script execution:  https://o365reports.com/2020/01/03/shared-mailbox-permission-report-to-csv/
============================================================================================
#>

#Accept input paramenters
param(
[switch]$FullAccess,
[switch]$SendAs,
[switch]$SendOnBehalf,
[string]$MBNamesFile,
[string]$UserName,
[string]$Password
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
  #Check for Exchange Online management module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell V2 module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
   Import-Module ExchangeOnlineManagement
  } 
  else 
  { 
   Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 
 Write-Host Connecting to Exchange Online...
 #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential
 }
 else
 {
  Connect-ExchangeOnline
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
  Write-Host ""
  Write-Host " Detailed report available in:" -NoNewline -ForegroundColor Yellow
  Write-Host $ExportCSV
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n 
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
