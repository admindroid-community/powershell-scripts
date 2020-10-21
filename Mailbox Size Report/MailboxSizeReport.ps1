Param
(
    [Parameter(Mandatory = $false)]
    [switch]$MFA,
    [switch]$SharedMBOnly,
    [switch]$UserMBOnly,
    [string]$MBNamesFile,
    [string]$UserName,
    [string]$Password
)

Function Get_MailboxSize
{
 $Stats=Get-MailboxStatistics -Identity $UPN
 $IsArchieved=$Stats.IsArchiveMailbox
 $ItemCount=$Stats.ItemCount
 $TotalItemSize=$Stats.TotalItemSize
 $TotalItemSizeinBytes= $TotalItemSize –replace “(.*\()|,| [a-z]*\)”, “”
 $TotalSize=$stats.TotalItemSize.value -replace "\(.*",""
 $DeletedItemCount=$Stats.DeletedItemCount
 $TotalDeletedItemSize=$Stats.TotalDeletedItemSize

 #Export result to csv
 $Result=@{'Display Name'=$DisplayName;'User Principal Name'=$upn;'Mailbox Type'=$MailboxType;'Primary SMTP Address'=$PrimarySMTPAddress;'IsArchieved'=$IsArchieved;'Item Count'=$ItemCount;'Total Size'=$TotalSize;'Total Size (Bytes)'=$TotalItemSizeinBytes;'Deleted Item Count'=$DeletedItemCount;'Deleted Item Size'=$TotalDeletedItemSize;'Issue Warning Quota'=$IssueWarningQuota;'Prohibit Send Quota'=$ProhibitSendQuota;'Prohibit send Receive Quota'=$ProhibitSendReceiveQuota}
 $Results= New-Object PSObject -Property $Result  
 $Results | Select-Object 'Display Name','User Principal Name','Mailbox Type','Primary SMTP Address','Item Count','Total Size','Total Size (Bytes)','IsArchieved','Deleted Item Count','Deleted Item Size','Issue Warning Quota','Prohibit Send Quota','Prohibit Send Receive Quota' | Export-Csv -Path $ExportCSV -Notype -Append 
}

Function main()
{
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

 #Connect Exchange Online with MFA
 if($MFA.IsPresent)
 {
  Connect-ExchangeOnline
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
  Connect-ExchangeOnline -Credential $Credential
 }

 #Output file declaration 
 $ExportCSV=".\MailboxSizeReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 

 $Result=""   
 $Results=@()  
 $MBCount=0
 $PrintedMBCount=0
 Write-Host Generating mailbox size report...
 
 #Check for input file
 if([string]$MBNamesFile -ne "") 
 { 
  #We have an input file, read it into memory 
  $Mailboxes=@()
  $Mailboxes=Import-Csv -Header "MBIdentity" $MBNamesFile
  foreach($item in $Mailboxes)
  {
   $MBDetails=Get-Mailbox -Identity $item.MBIdentity
   $UPN=$MBDetails.UserPrincipalName  
   $MailboxType=$MBDetails.RecipientTypeDetails
   $DisplayName=$MBDetails.DisplayName
   $PrimarySMTPAddress=$MBDetails.PrimarySMTPAddress
   $IssueWarningQuota=$MBDetails.IssueWarningQuota -replace "\(.*",""
   $ProhibitSendQuota=$MBDetails.ProhibitSendQuota -replace "\(.*",""
   $ProhibitSendReceiveQuota=$MBDetails.ProhibitSendReceiveQuota -replace "\(.*",""
   $MBCount++
   Write-Progress -Activity "`n     Processed mailbox count: $MBCount "`n"  Currently Processing: $DisplayName"
   Get_MailboxSize
   $PrintedMBCount++
  }
 }

 #Get all mailboxes from Office 365
 else
 {
  Get-Mailbox -ResultSize Unlimited | foreach {
   $UPN=$_.UserPrincipalName
   $Mailboxtype=$_.RecipientTypeDetails
   $DisplayName=$_.DisplayName
   $PrimarySMTPAddress=$_.PrimarySMTPAddress
   $IssueWarningQuota=$_.IssueWarningQuota -replace "\(.*",""
   $ProhibitSendQuota=$_.ProhibitSendQuota -replace "\(.*",""
   $ProhibitSendReceiveQuota=$_.ProhibitSendReceiveQuota -replace "\(.*",""
   $MBCount++
   Write-Progress -Activity "`n     Processed mailbox count: $MBCount "`n"  Currently Processing: $DisplayName"
   if($SharedMBOnly.IsPresent -and ($Mailboxtype -ne "SharedMailbox"))
   {
    return
   }
   if($UserMBOnly.IsPresent -and ($MailboxType -ne "UserMailbox"))
   {
    return
   }  
   Get_MailboxSize
   $PrintedMBCount++
  }
 }

 #Open output file after execution 
 If($PrintedMBCount -eq 0)
 {
  Write-Host No mailbox found
 }
 else
 {
  Write-Host `nThe output file contains $PrintedMBCount mailboxes.
  if((Test-Path -Path $ExportCSV) -eq "True") 
  {
   Write-Host `nThe Output file available in $ExportCSV -ForegroundColor Green
   $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
   {   
    Invoke-Item "$ExportCSV"   
   } 
  }
 }
 #Disconnect Exchange Online session
 Disconnect-ExchangeOnline -Confirm:$false | Out-Null
}
 . main

