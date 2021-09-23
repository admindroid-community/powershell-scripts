﻿<#
=============================================================================================
Name:           Audit Email Deletion in Office 365
Description:    This script exports deleted emails audit records to CSV
website:        o365reports.com
Script by:      O365Reports Team
For detailed Script execution: https://o365reports.com/2021/09/02/audit-email-deletion-in-office-365-mailbox-powershell
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$Mailbox,
    [string]$Subject,
    [string]$UserName,
    [string]$Password
)

Function Connect_Exo
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
}

$MaxStartDate=((Get-Date).AddDays(-89)).Date

#Getting anonymous activity for past 90 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
#Getting start date to generate anonymous activity report
While($true)
{
 if ($StartDate -eq $null)
 {
  $StartDate=Read-Host Enter start time for report generation '(Eg:04/28/2021)'
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
   Write-Host `nAudit can be retrieved only for past 90 days. Please select a date after $MaxStartDate -ForegroundColor Red
   return
  }
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}


#Getting end date to generate external sharing report
While($true)
{
 if ($EndDate -eq $null)
 {
  $EndDate=Read-Host Enter End time for report generation '(Eg: 04/28/2021)'
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

$OutputCSV=".\DeletedEmailsAuditReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
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
Write-Host `nRetrieving deleted emails from $StartDate to $EndDate...
$ProcessedAuditCount=0
$OutputEvents=0
$ExportResult=""   
$ExportResults=@()  
$RetriveOperation="SoftDelete,HardDelete,MoveToDeletedItems"

 #Checking whether the user is available
if(($Mailbox.Length -ne 0) -and ((Get-Mailbox -Identity $Mailbox) -eq $null))
{
 Write-Host Mailbox does not exist. Please check the mailbox name. -ForegroundColor Red
 exit
}

while($true)
{ 
 $ResultCount=0
 #Getting email deleted audit data for given time range
 Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $RetriveOperation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 | ForEach-Object {
  
  $ResultCount++
  $ProcessedAuditCount++
  Write-Progress -Activity "`n     Retrieving email deletion activities from $CurrentStart to $CurrentEnd.."`n" Processed audit record count: $ProcessedAuditCount"
  $MoreInfo=$_.auditdata
  $Operation=$_.Operations
  $AuditData=$_.auditdata | ConvertFrom-Json
  $TargetMailbox=$AuditData.MailboxOwnerUPN

  #Filter to view deleted emails from specific mailbox
  if(($mailbox.Length -ne 0) -and ($Mailbox -ne $TargetMailbox))
  {
   return
  }
  
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
   return
  } 

  $ActivityTime=Get-Date($AuditData.CreationTime)  #(Get-Date($AuditData.CreationTime)).ToLocalTime() Uncomment to view the Activity Time in local time
  $PerformedBy=$AuditData.userId
  $ResultStatus=$AuditData.ResultStatus
  $Folder=$AuditData.Folder.Path.Split("\")[1]
  $EventData=$AuditData.EventData
  
  #Export result to csv
  $OutputEvents++
  $ExportResult=@{'Activity Time'=$ActivityTime;'Target Mailbox'=$TargetMailbox;'Activity'=$Operation;'Performed By'=$PerformedBy;'No. of Emails Deleted'=$ConversationCount;'Email Subjects'=$Subjects;'Folder'=$Folder;'Result Status'=$ResultStatus;'Workload'=$Workload;'More Info'=$MoreInfo}
  $ExportResults= New-Object PSObject -Property $ExportResult  
  $ExportResults | Select-Object 'Activity Time','Activity','Target Mailbox','Performed By','No. of Emails Deleted','Email Subjects','Folder','Result Status','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
 }
 $currentResultCount=$CurrentResultCount+$ResultCount
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
}

#Open output file after execution
If($OutputEvents -eq 0)
{
 Write-Host No records found
}
else
{
 Write-Host `nThe output file contains $OutputEvents audit records
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host `nThe Output file availble in $OutputCSV -ForegroundColor Green
  Write-Host `nFor more Office 365 related PowerShell scripts, check https://o365reports.com -ForegroundColor Cyan
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