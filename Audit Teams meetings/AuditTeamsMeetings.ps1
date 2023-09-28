<#
=============================================================================================
Name:           Audit Microsoft Teams meetings using PowerShell
Description:    This script exports Teams meeting report and attendance report into 2 CSV files.
website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script uses modern authentication to connect to Exchange Online.   
2. The script can be executed with MFA enabled account too.   
3. Exports report results to CSV file.   
4. Allows you to generate a Teams meetings audit log.   
5. Exports Teams meeting attendance report. 
6. Helps to generate reports for custom periods. 
7. Automatically installs the EXO V2 and AzureAD module (if not installed already) upon your confirmation.  
8. The script is scheduler-friendly. I.e., Credential can be passed as a parameter instead of saving inside the script. 

For detailed Script execution:  https://o365reports.com/2021/12/08/microsoft-teams-meeting-attendance-report
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [switch]$NoMFA,
    [string]$UserName,
    [string]$Password
)

Function Connect_Modules
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
 #Check for Azure AD module
 $Module = Get-Module AzureAD -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Azure AD module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install the module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing AzureAD PowerShell module"
   Install-Module AzureAD -Repository PSGallery -AllowClobber -Force
  } 
  else 
  { 
   Write-Host AzureAD module is required to generate the report.Please install module using Install-Module AzureAD cmdlet. 
   Exit
  }
 }

 #Authentication using non-MFA
 if($NoMFA.IsPresent)
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
  Write-Host "Connecting Azure AD..."
  Connect-AzureAD -Credential $Credential | Out-Null
  Write-Host "Connecting Exchange Online PowerShell..."
  Connect-ExchangeOnline -Credential $Credential
 }
 #Connect to Exchange Online and AzureAD module using MFA 
 else
 {
  Write-Host "Connecting Exchange Online PowerShell..."
  Connect-ExchangeOnline
  Write-Host "Connecting Azure AD..."
  Connect-AzureAD | Out-Null
 }
}

Function Get_TeamMeetings
{
 $Result=""   
 $Results=@()  
 $Count=0
 Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "MeetingDetail" -ResultSize 5000 | ForEach-Object {
  $Count++
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
 Write-Host $Count meetings details are exported. -ForegroundColor Green
}

$MaxStartDate=((Get-Date).AddDays(-89)).Date
#Getting Teams meeting attendance report for past 90 days
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
  $StartDate=Read-Host Enter start time for report generation '(Eg:11/23/2021)'
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


#Getting end date for teams attendance report
While($true)
{
 if ($EndDate -eq $null)
 {
  $EndDate=Read-Host Enter End time for report generation '(Eg: 11/23/2021)'
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

Connect_Modules
Get_TeamMeetings
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
 Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $RetriveOperation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 | ForEach-Object {
  $ResultCount++
  $ProcessedAuditCount++
  Write-Progress -Activity "`n     Retrieving Team meeting participant report..."`n" Processed audit record count: $ProcessedAuditCount"
  $AuditData=$_.AuditData  | ConvertFrom-Json
  $MeetingID=$AuditData.MeetingDetailId
  $CreatedBy=$_.UserIDs
  $AttendeesInfo=($AuditData.Attendees)
  $Attendees=$AttendeesInfo.userObjectId
  $AttendeesType=$AttendeesInfo.RecipientType
  if($AttendeesType -ne "User")
  {
   $Attendees=$AttendeesInfo.DisplayName
  }
  else
  {
   $Attendees=(Get-AzureADUser -ObjectId $Attendees).UserPrincipalName
  }
  $JoinTime=(Get-Date($AuditData.JoinTime)).ToLocalTime()  #Get-Date($AuditData.JoinTime) Uncomment to view the Activity Time in UTC
  $LeaveTime=(Get-Date($AuditData.LeaveTime)).ToLocalTime()

   
  #Export result to csv
  $OutputEvents++
  $ExportResult=@{'Meeting id'=$MeetingID;'Created By'=$CreatedBy;'Attendees'=$Attendees;'Attendee Type'=$AttendeesType;'Joined Time'=$JoinTime;'Left Time'=$LeaveTime;'More Info'=$AuditData}
  $ExportResults= New-Object PSObject -Property $ExportResult  
  $ExportResults | Select-Object 'Meeting id','Created By','Attendees','Attendee Type','Joined Time','Left Time','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
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
 Write-Host `nThe Teams meeting attendance report contains $OutputEvents audit records -ForegroundColor Green
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host `n" The Teams meetings attendance report available in: " -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCSV"
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
  Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$OutputCSV"   
   Invoke-Item "$ExportCSV"
  } 
 }
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue


