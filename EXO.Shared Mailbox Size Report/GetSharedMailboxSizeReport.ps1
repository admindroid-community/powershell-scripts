<#
=============================================================================================
Name:           Export Shared Mailbox Size Report
Description:    This script exports Exchange Online shared mailbox size report to CSV
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.The script uses modern authentication to connect to Exchange Online.  
2.The script can be executed with MFA enabled account too.  
3.Exports report results to CSV. 
4.You can choose to either “export mailbox size of all the shared mailboxes” or pass an input file to “get usage statistics of specific shared mailboxes” alone. 
5.Automatically installs the EXO PowerShell module (if not installed already) upon your confirmation. 
6.The script is scheduler friendly. I.e., Credential can be passed as a parameter instead of saving inside the script. 
7.The script supports Certificate-based authentication (CBA).

For detailed Script execution: https://o365reports.com/2024/02/06/exchange-online-shared-mailbox-size-report-using-powershell/
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$MBNamesFile,
    [string]$UserName,
    [string]$Password,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Function Get_MailboxSize
{
 $Stats=Get-MailboxStatistics -Identity $UPN
 $ItemCount=$Stats.ItemCount
 $TotalItemSize=$Stats.TotalItemSize
 $TotalItemSizeinBytes= $TotalItemSize –replace “(.*\()|,| [a-z]*\)”, “”
 $TotalSize=$stats.TotalItemSize.value -replace "\(.*",""
 $DeletedItemCount=$Stats.DeletedItemCount
 $TotalDeletedItemSize=$Stats.TotalDeletedItemSize

 #Export result to csv
 $Result=@{'Shared Mailbox Name'=$DisplayName;'UPN'=$upn;'Mailbox Type'=$MailboxType;'Primary SMTP Address'=$PrimarySMTPAddress;'Archive Status'=$Archivestatus;'Item Count'=$ItemCount;'Total Size'=$TotalSize;'Total Size (Bytes)'=$TotalItemSizeinBytes;'Deleted Item Count'=$DeletedItemCount;'Deleted Item Size'=$TotalDeletedItemSize;'Issue Warning Quota'=$IssueWarningQuota;'Prohibit Send Quota'=$ProhibitSendQuota;'Prohibit send Receive Quota'=$ProhibitSendReceiveQuota}
 $Results= New-Object PSObject -Property $Result  
 $Results | Select-Object 'Shared Mailbox Name','UPN','Mailbox Type','Primary SMTP Address','Item Count','Total Size','Total Size (Bytes)','Archive Status','Deleted Item Count','Deleted Item Size','Issue Warning Quota','Prohibit Send Quota','Prohibit Send Receive Quota' | Export-Csv -Path $ExportCSV -Notype -Append 
}

Function main()
{
 #Check for EXO PowerShell module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
  } 
  else 
  { 
   Write-Host EXO PowerShell module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
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
 elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
   Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
 }
 else
 {
  Connect-ExchangeOnline
 }

 #Output file declaration 
 $ExportCSV=".\SharedMailboxSizeReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 

 $Result=""   
 $Results=@()  
 $MBCount=0
 $PrintedMBCount=0
 Write-Host Generating shared mailbox size report...
 
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
   #Check for archive enabled mailbox
   if(($MBDetails.ArchiveDatabase -eq $null) -and ($MBDetails.ArchiveDatabaseGuid -eq $MBDetails.ArchiveGuid))
   {
    $ArchiveStatus = "Disabled"
   }
   else
   {
    $ArchiveStatus= "Active"
   }
   $MBCount++
   Write-Progress -Activity "`n     Processed mailbox count: $MBCount "`n"  Currently Processing: $DisplayName"
   Get_MailboxSize
   $PrintedMBCount++
  }
 }

 #Get all mailboxes from Office 365
 else
 {
  Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | foreach {
   $UPN=$_.UserPrincipalName
   $Mailboxtype=$_.RecipientTypeDetails
   $DisplayName=$_.DisplayName
   $PrimarySMTPAddress=$_.PrimarySMTPAddress
   $IssueWarningQuota=$_.IssueWarningQuota -replace "\(.*",""
   $ProhibitSendQuota=$_.ProhibitSendQuota -replace "\(.*",""
   $ProhibitSendReceiveQuota=$_.ProhibitSendReceiveQuota -replace "\(.*",""
   $MBCount++
   Write-Progress -Activity "`n     Processed mailbox count: $MBCount "`n"  Currently Processing: $DisplayName"
 
   #Check for archive enabled mailbox
   if(($_.ArchiveDatabase -eq $null) -and ($_.ArchiveDatabaseGuid -eq $_.ArchiveGuid))
   {
    $ArchiveStatus = "Disabled"
   }
   else
   {
    $ArchiveStatus= "Active"
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

   Write-Host `n The Output file available in: -NoNewline -ForegroundColor Yellow
   Write-Host $ExportCSV 
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
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 
}
 . main 
 

