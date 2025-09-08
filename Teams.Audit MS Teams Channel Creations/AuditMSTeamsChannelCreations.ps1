<#
=============================================================================================
Name:           Audit Microsoft Teams Channel creations in Microsoft 365
Version:        1.0
Website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script uses modern authentication to retrieve audit logs.     
2. The script can be executed with an MFA enabled account too.     
3. The script retrieves audit log for 180 days, by default.
4. Allows you generate a Channel creation audit report for a custom period. 
5. Helps to find recently created channels. e.g., MS Teams channels created in the last 30 days. 
6. Exports report results to CSV file.     
7. Identifies Channels created by external/guest users. 
8. Helps to monitor channels created by a specific user.
9. Helps to get private channels creations alone
10.Tracks shared channels creations alone.
11.Findsind standard channels creations alone.   
12. Automatically installs the EXO module (if not installed already) upon your confirmation.   
13. The script is scheduler friendly. i.e., Credentials can be passed as a parameter instead of saved inside the script. 
14. The script supports Certificate-based Authentication (CBA) too.

For detailed script execution: https://o365reports.com/2023/12/19/audit-ms-teams-channel-creations-using-powershell/ 
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [Int]$RecentlyCreatedChannel_In_Days,
    [switch]$PrivateChannelsOnly,
    [switch]$SharedChannelsOnly,
    [switch]$StandardChannelsOnly,
    [Switch]$CreatedByExternalUsersOnly,
    [String]$CreatedBy,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password
)

Function Connect_Exo
{
 #Check for EXO module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
   Import-Module ExchangeOnlineManagement
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
}

$MaxStartDate=((Get-Date).AddDays(-179)).Date
if($RecentlyCreatedChannel_In_Days -ne "")
{
 $StartDate=((Get-Date).AddDays(-$RecentlyCreatedChannel_In_Days)).Date
 $EndDate=(Get-Date).Date
}

#Audit MS Teams channels creations in the past 180 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
#Getting start date to audit MS Teams channel creation report
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


#Getting end date to audit MS Teams channel creations
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

$OutputCSV=".\Audit_MSTeams_Channel_Creations_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
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

#Connect_EXO
$CurrentResultCount=0
$AggregateResultCount=0
Write-Host `nAuditing Microsoft Teams channel creations from $StartDate to $EndDate... -ForegroundColor Cyan
$ProcessedAuditCount=0
$OutputEvents=0
$ExportResult=""   
$ExportResults=@()  
$Operations="ChannelAdded"

if($CreatedByExternalUsersOnly.IsPresent)
{
 $UserIds="*#EXT*"
}
elseif($CreatedBy -ne "")
{
 $UserIds=$CreatedBy
}
else
{
 $UserIds="*"
}

while($true)
{ 
 #Getting audit data for the given time range
 Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $Operations -UserIds $UserIds -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 | ForEach-Object {
  $ResultCount++
  $ProcessedAuditCount++
  $PrintFlag=$true
  Write-Progress -Activity "`n     Retrieving Microsoft Teams Channel creations from $CurrentStart to $CurrentEnd.."`n" Processed audit record count: $ProcessedAuditCount"
  $MoreInfo=$_.auditdata
  $Operation=$_.Operations
  $CreatedBy=$_.UserIds
  $AuditData=$_.auditdata | ConvertFrom-Json
  $ResultStatus=$AuditData.ResultStatus
  $EventTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()  #Get-Date($AuditData.CreationTime) Uncomment to view the Activity Time in UTC
  $ChannelName=$AuditData.ChannelName
  $TeamName=$AuditData.TeamName
  $ChannelType=$AuditData.ChannelType

  #Filter private channels
  if(($PrivateChannelsOnly.IsPresent) -and ($ChannelType -ne "Private"))
  {
   return
  }

  #Filter shared channel creations

  if(($SharedChannelsOnly.IsPresent) -and ($ChannelType -ne "Shared"))
  {
   return
  }

  if(($StandardChannelsOnly.IsPresent) -and ($ChannelType -ne "Standard"))
  {
   return
  }



  #Export result to csv
  if($PrintFlag -eq "True")
  {
   $OutputEvents++
   $ExportResult=@{'Creation Time'=$EventTime;'Created By'=$CreatedBy; 'Channel Name'=$ChannelName;'Team Name'=$TeamName;'Channel Type'=$ChannelType;'More Info'=$MoreInfo}
   $ExportResults= New-Object PSObject -Property $ExportResult  
   $ExportResults | Select-Object 'Creation Time','Channel Name','Created By','Channel Type','Team Name','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
  }
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
}
else
{
 Write-Host `nThe output file contains $OutputEvents audit records
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host `n "The Output file availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCSV" `n 
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$OutputCSV"   
  } 
 }
}
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
#Disconnect Exchange Online session
#Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue