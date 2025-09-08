<#
=============================================================================================
Name:           Audit Microsoft Teams meetings and export teams meeting attendance report using PowerShell
Version:        2.0
Description:    This script exports Teams meeting report and attendance report into 2 CSV files.
website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. Exports a list of Teams meetings.   
2. Exports Teams meeting attendance report.   
3. Exports report results to CSV file.   
4. The script can be executed with MFA enabled account too. 
5. Supports certificate-based authentication (CBA) too.
6. Helps to generate reports for custom periods. 
7. Automatically installs the EXO PowerShell module (if not installed already) upon your confirmation.  
8. The script is scheduler-friendly.  

For detailed Script execution:  https://o365reports.com/2021/12/08/microsoft-teams-meeting-attendance-report

Change Log:
~~~~~~~~~~~

    V1.0 (Dec 09, 2021) - File created
    V1.1 (Sep 28, 2023) - Minor usability improvements.
    V2.0 (Mar 05, 2025) - Added certificate-based authentication support to enhance scheduling capability. 
                          Removed Azure AD PowerShell module dependency.
                          Increased audit log retrieval range from 90 to 180 days
                          Usability improvements
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserName,
    [string]$Password,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

#Getting StartDate and EndDate for Audit log
if ((($StartDate -eq $null) -and ($EndDate -ne $null)) -or (($StartDate -ne $null) -and ($EndDate -eq $null)))
{
 Write-Host `nPlease enter both StartDate and EndDate for Audit log collection -ForegroundColor Red
 exit
}   
elseif(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $StartDate=(((Get-Date).AddDays(-180))).Date
 $EndDate=Get-Date
}
else
{
 $StartDate=[DateTime]$StartDate
 $EndDate=[DateTime]$EndDate
 if($StartDate -lt ((Get-Date).AddDays(-180)))
 { 
  Write-Host `nAudit log can be retrieved only for past 180 days. Please select a date after (Get-Date).AddDays(-180) -ForegroundColor Red
  Exit
 }
 if($EndDate -lt ($StartDate))
 {
  Write-Host `nEnd time should be later than start time -ForegroundColor Red
  Exit
 }
}

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
 if(($AdminName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $AdminName,$SecuredPassword
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

Function Get_TeamMeetings
{
 $Result=""   
 $Results=@()  
 Write-Host `nRetrieving Teams meeting details from $StartDate to $EndDate...
 Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "MeetingDetail" -ResultSize 5000 | ForEach-Object {
  $Global:MeetingCount++
  Write-Progress -Activity "`n     Retrieving Teams meetings data from $StartDate to $EndDate.."`n" Processed Teams meetings count: $Count"
  $AuditData=$_.AuditData  | ConvertFrom-Json
  $MeetingID=$AuditData.ID
  $CreatedBy=$AuditData.UserId
  $StartTime=(Get-Date($AuditData.StartTime)).ToLocalTime()
  $EndTime=(Get-Date($AuditData.EndTime)).ToLocalTime()
  $MeetingURL=$AuditData.MeetingURL
  $MeetingType=$AuditData.ItemName
  $Result=@{'Meeting id'=$MeetingID;'Created By'=$CreatedBy;'Start Time'=$Star=$StartTime;'End Time'=$EndTime;'Meeting Type'=$MeetingType;'Meeting Link'=$MeetingURL;'More Info'=$AuditData}
  $Results= New-Object PSObject -Property $Result  
  $Results | Select-Object 'Meeting id','Created By','Start Time','End Time','Meeting Type','Meeting Link','More Info' | Export-Csv -Path $ExportCSV -Notype -Append 
 }
 if($MeetingCount -ne 0)
 {
  Write-Host $Global:MeetingCount meetings details are exported. 
 }
 else
 {
  Write-Host "No meetings found"
 }
}

$MaxStartDate=((Get-Date).AddDays(-180)).Date
#Getting Teams meeting attendance report for past 180 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date)#.Date
 $StartDate=$MaxStartDate
}
#Getting start date to generate Teams meetings attendance report
While($true)
{
 if ($StartDate -eq $null)
 {
  $StartDate=Read-Host Enter start time for report generation '(Eg:11/23/2024)'
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
   Write-Host `nAudit can be retrieved only for past 180 days. Please select a date after $MaxStartDate -ForegroundColor Red
   return
  }
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}


#Getting end date for teams attendance report
While($true)
{
 if ($EndDate -eq $null)
 {
  $EndDate=Read-Host Enter End time for report generation '(Eg: 11/23/2024)'
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

$OutputCSV=".\TeamsMeetingAttendanceReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$ExportCSV=".\TeamsMeetingsReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
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
$Global:MeetingCount=0
Get_TeamMeetings

#Get participants details if any meetings found
if($Global:MeetingCount -ne 0)
{
 $CurrentResultCount=0
 $AggregateResultCount=0
 Write-Host `nGenerating Teams meeting attendance report from $StartDate to $EndDate...
 $ProcessedAuditCount=0
 $OutputEvents=0
 $ExportResult=""   
 $ExportResults=@()  
 $RetriveOperation="MeetingParticipantDetail"
 while($true)
 { 
  #Getting Teams meeting participants report for the given time range
  $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $RetriveOperation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 
  $ResultsCount=($Results|Measure-Object).count 
  foreach($Result in $Results)
  {
   $ProcessedAuditCount++
   $AuditData=$Result.AuditData  | ConvertFrom-Json
   $MeetingID=$AuditData.MeetingDetailId
   $CreatedBy=$Result.UserIDs
   $AttendeesInfo=($AuditData.Attendees)
   $Attendees=$AttendeesInfo.userObjectId
   $AttendeesType=$AttendeesInfo.RecipientType
   if($AttendeesType -ne "User")
   {
    $Attendees=$AttendeesInfo.DisplayName
   }
   else
   {
    $Attendees=(Get-ExoRecipient -Identity $Attendees).DisplayName
   }
   $JoinTime=(Get-Date($AuditData.JoinTime)).ToLocalTime()  #Get-Date($AuditData.JoinTime) Uncomment to view the Activity Time in UTC
   $LeaveTime=(Get-Date($AuditData.LeaveTime)).ToLocalTime()

   
   #Export result to csv
   $OutputEvents++
   $ExportResult=@{'Meeting id'=$MeetingID;'Created By'=$CreatedBy;'Attendees'=$Attendees;'Attendee Type'=$AttendeesType;'Joined Time'=$JoinTime;'Left Time'=$LeaveTime;'More Info'=$AuditData}
   $ExportResults= New-Object PSObject -Property $ExportResult  
   $ExportResults | Select-Object 'Meeting id','Created By','Attendees','Attendee Type','Joined Time','Left Time','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
  }
  $currentResultCount=$CurrentResultCount+$ResultsCount
  Write-Progress -Activity "`n     Retrieving audit log for $CurrentStart : $CurrentResultCount records"`n" Total processed audit record count: $ProcessedAuditCount"
  if(($CurrentResultCount -eq 50000) -or ($ResultsCount -lt 5000))
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
}

#Open output file after execution
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
 Write-Host " to get access to 1900+ Microsoft 365 reports. ~~" -ForegroundColor Green `n
 
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host `nThe Teams meeting attendance report contains $OutputEvents audit records
  Write-Host `n" The Teams meetings attendance report available in: " -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCSV"
 
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$OutputCSV"   
   Invoke-Item "$ExportCSV"
  } 
 }


#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue