<#
=============================================================================================
Name:           Find who sent email from shared mailbox
Description:    This script finds who sent emails from shared mailbox
Version:        3.0
Website:        o365reports.com

Change Log
~~~~~~~~~~

    V1 (11/5/22) - Initial version 
    V2 (9/28/23) - Minor changes
    V2 (9/18/24) - Added support for certificate-based authentication and extended audit log retrieval perid from 90 to 180 days 


Script Highlights: 
~~~~~~~~~~~~~~~~~~

1.Helps to generate audit reports for custom periods.  
2.Tracks email sent activities from a specific shared mailbox. 
3.Allows to audit send as activities separately. 
4.Allows to track send on behalf activities separately. The script uses modern authentication to retrieve audit logs.    
5.Exports report results to CSV file.   
6.Automatically installs the EXO module (if not installed already) upon your confirmation. 
7.The script can be executed with an MFA-enabled account too.    
8.Supports Certificate-based Authentication (CBA) too
9.The script is scheduler-friendly.

For detailed script execution: https://o365reports.com/2022/05/11/find-who-sent-email-from-shared-mailbox-in-office-365-using-powershell
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$SharedMBIdentity,
    [switch]$SendAsOnly,
    [Switch]$SendOnBehalfOnly,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password
)

$MaxStartDate=((Get-Date).AddDays(-179)).Date


#Retrive audit log for the past 180 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
#Getting start date to audit export report
While($true)
{
 if ($StartDate -eq $null)
 {
  $StartDate=Read-Host Enter start time for report generation '(Eg:12/15/2023)'
 }
 Try
 {
  $Date=[DateTime]$StartDate
  if($Date -ge $MaxStartDate)
  { 
   break
  }
  else
  {
   Write-Host `nAudit can be retrieved only for the past 180 days. Please select a date after $MaxStartDate -ForegroundColor Red
   return
  }
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}


#Getting end date to export audit report
While($true)
{
 if ($EndDate -eq $null)
 {
  $EndDate=Read-Host Enter End time for report generation '(Eg: 12/15/2023)'
 }
 Try
 {
  $Date=[DateTime]$EndDate
  if($EndDate -lt ($StartDate))
  {
   Write-Host End time should be later than start time -ForegroundColor Red
   return
  }
  break
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}


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

$Location=Get-Location
$OutputCSV="$Location\AuditWhoSentEmailsFromSharedMailbox_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Check whether CurrentEnd exceeds EndDate
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

if($CurrentStart -eq $CurrentEnd)
{
 Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
 Exit
}

Connect_EXO
$CurrentResultCount=0
$AggregateResultCount=0
Write-Host `nAuditing emails sent from shared mailboxes - $StartDate to $EndDate...
$ProcessedAuditCount=0
$OutputEvents=0
$ExportResult=""   
$ExportResults=@()  

#Check whether to retrieve for all the Shared mailboxes or a specific mailbox
if($SharedMBIdentity -eq "")
{
 $SMBs=(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox).PrimarySMTPAddress
}
else
{
 #Checking whether the user is available
 if((Get-Mailbox -Identity $SharedMBIdentity -RecipientTypeDetails Sharedmailbox) -eq $null)
 {
  Write-Host Given Shared Mailbox does not exist. Please check the name. -ForegroundColor Red
  exit
 }
}


#Check for SendAs and SendOnBehalf filter
if($SendAsOnly.IsPresent)
{
 $Operations="SendAs"
}
elseif($SendOnBehalfOnly.IsPresent)
{
 $Operations="SendOnBehalf"
}
else
{
 $Operations="SendAs,SendOnBehalf"
}

while($true)
{
 #Getting audit data for the given time range
 Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $Operations -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 | ForEach-Object {
  $ResultCount++
  $ProcessedAuditCount++
  Write-Progress -Activity "`n     Retrieving email sent activities from $CurrentStart to $CurrentEnd.."`n" Processed audit record count: $ProcessedAuditCount"
  $MoreInfo=$_.auditdata
  $Operation=$_.Operations
  $SentBy=$_.UserIds
  if($SentBy -eq "S-1-5-18")
  {
   $SentBy="Service account"
  }
  $AuditData=$_.auditdata | ConvertFrom-Json
  If($Operation -eq "SendAs")
  {
   $SentFrom=$AuditData.SendAsUserSMTP
  }
  else
  {
   $SentFrom=$AuditData.SendOnBehalfOfUsersmtp
  }

  if($SharedMBIdentity -eq "")
  {
   if($SMBs -notcontains $SentFrom)
   {
    return
   }
  }
  elseif($SentFrom -ne $SharedMBIdentity)
  {
   return
  }
   
  $Subject=$AuditData.Item.Subject
  $Result=$AuditData.ResultStatus
  $PrintFlag="True"
  $SentTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()  #Get-Date($AuditData.CreationTime) Uncomment to view the Activity Time in UTC
   #Export result to csv
   $OutputEvents++
     #Export result to csv
  $ExportResult=@{'Sent Time'=$SentTime;'Sent By'=$SentBy;'Sent From'=$SentFrom; 'Subject'=$Subject; 'Operation'=$Operation;'Result'=$Result;'More Info'=$MoreInfo}
  $ExportResults= New-Object PSObject -Property $ExportResult  
  $ExportResults | Select-Object 'Sent Time','Sent By','Sent From','Subject','Operation','Result','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
  
 }
 $currentResultCount=$currentResultCount+$ResultCount
 
 if($CurrentResultCount -ge 50000)
 {
  Write-Host Retrieved max record for current range.Proceeding further may cause data loss or rerun the script with reduced time interval. -ForegroundColor Red
  $Confirm=Read-Host `nAre you sure you want to continue? [Y] Yes [N] No
  if($Confirm -match "[Y]")
  {
   Write-Host Proceeding audit log collection with data loss
   [DateTime]$CurrentStart=$CurrentEnd
   [DateTime]$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
   $CurrentResultCount=0
   if($CurrentEnd -gt $EndDate)
   {
    $CurrentEnd=$EndDate
   }
  }
  else
  {
   Write-Host Please rerun the script with reduced time interval -ForegroundColor Red
   Exit
  }
 }

 
 if($ResultCount -lt 5000)
 { 
  if($CurrentEnd -eq $EndDate)
  {
   break
  }
  $CurrentStart=$CurrentEnd 
  if($CurrentStart -gt (Get-Date))
  {
   break
  }
  $CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
  $CurrentResultCount=0
  if($CurrentEnd -gt $EndDate)
  {
   $CurrentEnd=$EndDate
  }
 }                                                                                             
 $ResultCount=0
}

#Open output file after execution
If($OutputEvents -eq 0)
{
 Write-Host No records found

 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n 
}
else
{
 Write-Host `nThe output file contains $OutputEvents audit records `n
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n 
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host " The Output file availble in:" -NoNewline -ForegroundColor Yellow 
  Write-Host $OutputCSV 
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$OutputCSV"   
  } 
 }
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
  