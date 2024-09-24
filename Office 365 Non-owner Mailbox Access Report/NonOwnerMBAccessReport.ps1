<#
=============================================================================================
Name:           Export Non-Owner Mailbox Access Report
Version:        2.0
Website:        o365reports.com


Script Highlights: 
~~~~~~~~~~~~~~~~~~

1.Allows you to filter out external users’ access. 
2.The script can be executed with MFA enabled account too. 
3.Exports the report to CSV 
4.This script is scheduler friendly. I.e., credentials can be passed as a parameter instead of saving inside the script. 
5.You can narrow down the audit search for a specific date range. 
6.The script supports Certificate-based authentication too.

For detailed script execution:  https://o365reports.com/2020/02/04/export-non-owner-mailbox-access-report-to-csv/



Change Log
~~~~~~~~~~

    V1.0 (Feb 17, 2020) - File created
    V1.1 (Oct 06, 2023) - Minor changes
    V2.0 (Nov 25, 2023) - Added certificate-based authentication support to enhance scheduling capability
    V2.1 (Sep 24, 2024) - Special handling done to track send as and send-on-behalf activities accurately.
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [Boolean]$IncludeExternalAccess=$false,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password
)

#Getting StartDate and EndDate for Audit log
if ((($StartDate -eq $null) -and ($EndDate -ne $null)) -or (($StartDate -ne $null) -and ($EndDate -eq $null)))
{
 Write-Host `nPlease enter both StartDate and EndDate for Audit log collection -ForegroundColor Red
 exit
}
elseif(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $StartDate=(((Get-Date).AddDays(-90))).Date
 $EndDate=Get-Date
}
else
{
 $StartDate=[DateTime]$StartDate
 $EndDate=[DateTime]$EndDate
 if($StartDate -lt ((Get-Date).AddDays(-90)))
 {
  Write-Host `nAudit log can be retrieved only for past 90 days. Please select a date after (Get-Date).AddDays(-90) -ForegroundColor Red
  Exit
 }
 if($EndDate -lt ($StartDate))
 {
  Write-Host `nEnd time should be later than start time -ForegroundColor Red
  Exit
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

$OutputCSV=".\NonOwner-Mailbox-Access-Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
$Operation='ApplyRecord','Copy','Create','FolderBind','HardDelete','MessageBind','Move','MoveToDeletedItems','RecordDelete','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules'


#Check whether CurrentEnd exceeds EndDate(checks for 1st iteration)
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

$AggregateResults = 0
$CurrentResult= @()
$CurrentResultCount=0
$NonOwnerAccess=0
Write-Host `nRetrieving audit log from $StartDate to $EndDate... -ForegroundColor Yellow

while($true)
{
 #Write-Host Retrieving audit log between StartDate $CurrentStart to EndDate $CurrentEnd ******* IntervalTime $IntervalTimeInMinutes minutes
 if($CurrentStart -eq $CurrentEnd)
 {
  Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
  Exit
 }


 #Getting Non-Owner mailbox access for a given time range
 else
 {
  $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $Operation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 }
 $AllAuditData=@()
 $AllAudits=
 foreach($Result in $Results)
 {
  $AuditData=$Result.auditdata | ConvertFrom-Json

  #Remove owner access
  if($AuditData.LogonType -eq 0)
  {
   continue
  }

  #Filter for external access
  if(($IncludeExternalAccess -eq $false) -and ($AuditData.ExternalAccess -eq $true))
  {
   continue
  }

  #Processing non-owner mailbox access records
  if(($AuditData.LogonUserSId -ne $AuditData.MailboxOwnerSid) -or ((($AuditData.Operation -eq "SendAs") -or ($AuditData.Operation -eq "SendOnBehalf")) -and ($AuditData.UserType -eq 0)))
  {
   $AuditData.CreationTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()
   if($AuditData.LogonType -eq 1)
   {
    $LogonType="Administrator"
   }
   elseif($AuditData.LogonType -eq 2)
   {
    $LogonType="Delegated"
   }
   else
   {
    $LogonType="Microsoft datacenter"
   }
   if($AuditData.Operation -eq "SendAs")
   {
    $AccessedMB=$AuditData.SendAsUserSMTP
    $AccessedBy=$AuditData.UserId
   }
   elseif($AuditData.Operation -eq "SendOnBehalf")
   {
    $AccessedMB=$AuditData.SendOnBehalfOfUserSmtp
    $AccessedBy=$AuditData.UserId
   }
   else
   {
    $AccessedMB=$AuditData.MailboxOwnerUPN
    $AccessedBy=$AuditData.UserId
   }
   if($AccessedMB -eq $AccessedBy)
   {
    Continue
   }
  $NonOwnerAccess++
  $AllAudits=@{'Access Time'=$AuditData.CreationTime;'Accessed by'=$AccessedBy;'Performed Operation'=$AuditData.Operation;'Accessed Mailbox'=$AccessedMB;'Logon Type'=$LogonType;'Result Status'=$AuditData.ResultStatus;'External Access'=$AuditData.ExternalAccess;'More Info'=$Result.auditdata}
  $AllAuditData= New-Object PSObject -Property $AllAudits
  $AllAuditData | Sort 'Access Time','Accessed by' | select 'Access Time','Logon Type','Accessed by','Performed Operation','Accessed Mailbox','Result Status','External Access','More Info' | Export-Csv $OutputCSV -NoTypeInformation -Append
 }
 }
  #$CurrentResult += $Results
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
#Open output file after execution
else
{
 Write-Host `nThe output file contains $NonOwnerAccess audit records
 if((Test-Path -Path $OutputCSV) -eq "True")
 {
  Write-Host `nThe Output file available in: -NoNewline -ForegroundColor Yellow
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
Disconnect-ExchangeOnline -Confirm:$false