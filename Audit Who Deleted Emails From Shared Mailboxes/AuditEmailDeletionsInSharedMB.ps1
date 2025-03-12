<#
=============================================================================================
Name: Audit Email Deletions in shared mailbox using PowerShell
Version: 1.0 
Website: o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the Exchange Online PowerShell module (if not installed already) upon your confirmation. 
2. Allows to track all shared mailbox email deletions.
3. Tracks email deletions from a specific shared mailbox.
4. Allows you to generate a shared mailbox email deletion audit report for a custom period.
5. The script can be executed with MFA enabled account too.
6. Exports a specific user’s email deletions activities in shared mailboxes.
7. Filters deleted emails from shared mailbox by its subject. 
8. It exports the report result to CSV format.
9. It can be executed with Certificate-based Authentication (CBA) too.
10. The script is scheduler friendly.

For detailed script execution:https://o365reports.com/2025/03/11/audit-email-deletions-in-microsoft-365-shared-mailboxes-powershell/ 
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$SharedMBIdentity,
    [string]$UserId,
    [string]$Subject,
    [string]$UserName,
    [string]$Password,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint
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
  $StartDate=Read-Host Enter start time for report generation '(Eg:09/24/2024)'
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
  $EndDate=Read-Host Enter End time for report generation '(Eg: 9/24/2024)'
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
Connect_Exo


$OutputCSV=".\DeletedEmailsAuditReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Check whether CurrentEnd exceeds EndDate
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}


$CurrentResultCount=0
$AggregateResultCount=0
Write-Host `nRetrieving deleted emails from $StartDate to $EndDate...
$ProcessedAuditCount=0
$OutputEvents=0
$ExportResult=""   
$ExportResults=@()  
$RetriveOperation="SoftDelete,HardDelete,MoveToDeletedItems"

#Check whether to retrieve for all the Shared mailboxes or a specific mailbox
if($SharedMBIdentity -eq "")
{
 $SMBs=(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox).PrimarySMTPAddress
}
else
{
 #Checking whether the shared mailbox is available
 if((Get-Mailbox -Identity $SharedMBIdentity -RecipientTypeDetails Sharedmailbox) -eq $null)
 {
  Write-Host Given Shared Mailbox does not exist. Please check the name. -ForegroundColor Red
  exit
 }
}

while($true)
{ 
 $ResultCount=0
 #Getting email deleted audit data for given time range
 if($UserId.Length -ne 0)
 { 
  $Results=Search-UnifiedAuditLog -UserIds $UserId -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $RetriveOperation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 
 }
 else
 { 
  $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $RetriveOperation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 
 }
 $ResultCount=($Results | Measure-Object).count
 foreach($Result in $Results)
 {
  $ProcessedAuditCount++
  Write-Progress -Activity "`n     Retrieving email deletion activities from $CurrentStart to $CurrentEnd.."`n" Processed audit record count: $ProcessedAuditCount"
  $MoreInfo=$Result.auditdata
  $Operation=$Result.Operations
  $AuditData=$Result.auditdata | ConvertFrom-Json
  $TargetMailbox=$AuditData.MailboxOwnerUPN

 
  $EmailSubject=$AuditData.AffectedItems.Subject
  if($EmailSubject -eq $null)
  {
   $EmailSubject="-"
  }
  $ConversationCount=$EmailSubject.count
  $Subjects=$EmailSubject -join ", "

  #Filter emails by subject
  if(($Subject.Length -ne 0) -and ($Subjects -notmatch $Subject))
  { 
   continue
  } 

  if($SharedMBIdentity -eq "")
  {
   if($SMBs -notcontains $TargetMailbox)
   {
    continue
   }
  }
  elseif($TargetMailbox -ne $SharedMBIdentity)
  {
   continue
  }

  $ActivityTime=Get-Date($AuditData.CreationTime)  #(Get-Date($AuditData.CreationTime)).ToLocalTime() Uncomment to view the Activity Time in local time
  $PerformedBy=$AuditData.userId
  $ResultStatus=$AuditData.ResultStatus
  $Folder=$AuditData.Folder.Path.Split("\")[1]
  $EventData=$AuditData.EventData
  
  #Export result to csv
  $OutputEvents++
  $ExportResult=@{'Activity Time'=$ActivityTime;'Shared Mailbox Name'=$TargetMailbox;'Activity'=$Operation;'Performed By'=$PerformedBy;'No. of Emails Deleted'=$ConversationCount;'Email Subjects'=$Subjects;'Folder'=$Folder;'Result Status'=$ResultStatus;'Workload'=$Workload;'More Info'=$MoreInfo}
  $ExportResults= New-Object PSObject -Property $ExportResult  
  $ExportResults | Select-Object 'Activity Time','Shared Mailbox name','Performed By','Activity','No. of Emails Deleted','Email Subjects','Folder','Result Status','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
 }

  $currentResultCount=$CurrentResultCount+($Results.count)
 $AggregateResults +=$Results.count
 Write-Progress -Activity "`n     Retrieving audit log for $CurrentStart : $CurrentResultCount records"`n" Total processed audit record count: $AggregateResults"
 if(($CurrentResultCount -eq 50000) -or ($Results.count -lt 5000))
 {
  if($CurrentResultCount -eq 50000)
  {
   Write-Host Retrieved max record for the current range.Proceeding further may cause data loss or rerun the script with reduced time interval. -ForegroundColor Red
   $Confirm=Read-Host `nAre you sure you want to continue? [Y] Yes [N] No
   if($Confirm -notmatch "[Y]")
   {
    Write-Host Please rerun the script with reduced time interval -ForegroundColor Red
    Exit
   }
   else
   {
    Write-Host Proceeding audit log collection with data loss
   }
  }
  #Check for last iteration
  if(($CurrentEnd -eq $EndDate))
  {
   break
  }
  [DateTime]$CurrentStart=$CurrentEnd
  #Break loop if start date exceeds current date(There will be no data)
  if($CurrentStart -gt (Get-Date))
  {
   break
  }
  [DateTime]$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
  if($CurrentEnd -gt $EndDate)
  {
   $CurrentEnd=$EndDate
  }

  $CurrentResultCount=0
  $CurrentResult = @()
 }
}
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  
If($AggregateResults -eq 0)
{
 Write-Host No records found
}


else
{
 Write-Host `nThe output file contains $OutputEvents audit records
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host `n" The Output file available in:"  -NoNewline -ForegroundColor Yellow; Write-Host $OutputCSV 
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