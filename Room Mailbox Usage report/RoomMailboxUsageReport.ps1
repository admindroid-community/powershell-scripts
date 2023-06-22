<#
=============================================================================================
Name:           Get Office 365 Room Mailbox Usage Statistics Using PowerShell
Description:    This script gives detailed information on all Office 365 room mailboxes usage
Version:        2.0
Website:        o365reports.com

Script Highlights:
~~~~~~~~~~~~~~~~~
1.Automatically installs the MS Graph PowerShell module upon your confirmation when it is not available on your machine. 
2.Also, you can execute this script with certificate-based authentication (CBA). 
3.You can execute the script with an MFA-enabled account too. 
4.Helps to filter details about online meetings alone.
5.Gets meetings details by organizers.
6.Helps to identify meeting scheduled for today.
7.Further, this script is scheduler-friendly! Therefore, you can automate the report generation easily.
8.The script generates 2 output file, one with detailed info and another with summary info.

For detailed script execution: https://o365reports.com/2023/05/23/get-office-365-room-mailbox-usage-statistics-using-powershell/
============================================================================================
#>Param
(
    [switch]$OnlineMeetingOnly,
    [switch]$ShowTodaysMeetingsOnly,
    [String]$OrgEmailId,
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Function Connect_MgGraph
{
 #Check for module installation
 $Module=Get-Module -Name microsoft.graph -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Microsoft Graph PowerShell SDK is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Microsoft Graph PowerShell module..."
   Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
  }
  else
  {
   Write-Host "Microsoft Graph PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph cmdlet." 
   Exit
  }
 }
 #Disconnect Existing MgGraph session
 if($CreateSession.IsPresent)
 {
  Disconnect-MgGraph
 }

 #Connecting to MgGraph beta
 Select-MgProfile -Name beta
 Write-Host Connecting to Microsoft Graph...
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint 
 }
 else
 {
  Connect-MgGraph -Scopes "Place.Read.All","User.Read.All","Calendars.Read","Calendars.Read.Shared"  
 }
}
Connect_MgGraph

$ExportCSV = ".\RoomMailboxUsageReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportSummaryCSV=".\RoomMailboxUsageSummaryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult=""   
$ExportSummary=""
$startDate=(Get-date).AddDays(-30).Date
$EndDate=(Get-date).AddDays(1).Date
$MbCount=0
$PrintedMeetings=0

#Retrieving all room mailboxes
Get-MgPlaceAsRoom -All | foreach {
 $RoomAddress=$_.EmailAddress
 $RoomName=$_.DisplayName
 $MeetingCount=0
 $Count++
 $RoomUsage=0
 $OnlineMeetingCount=0
 $AllDayMeetingCount=0
 
 Get-MgUserCalendarView  -UserId $RoomAddress -StartDateTime $startDate -EndDateTime $EndDate -All | foreach {
  Write-Progress -Activity "`n     Processing room: $Count - $RoomAddress : Meeting Count - $MeetingCount"
  if($_.IsCancelled -eq $false)
  {
   $Print=1
   $MeetingCount++
   $Organizer=$_.Organizer.EmailAddress.Address
   $MeetingSubject=$_.Subject
   $IsAllDayMeeting=$_.IsAllDay
   $IsOnlineMeeting=$_.IsOnlineMeeting
   if($IsOnlineMeeting -eq $true)
   {
    $OnlineMeetingCount++
   }
   if($IsAllDayMeeting -eq $true)
   {
    $AllDayMeetingCount++
   }
   $MeetingStartTimeZone=$_.OriginalStartTimeZone
   $MeetingCreatedTime=$_.CreatedDateTime
   $MeetingLastModifiedTime=$_.LastModifiedDateTime
   [Datetime]$MeetingStart=$_.Start.DateTime
   $MeetingStartTime=$MeetingStart.ToLocalTime()
   [Datetime]$MeetingEnd=$_.End.DateTime
   $MeetingEndTime=$MeetingEnd.ToLocalTime()
   if($_.IsAllDay -eq $true)
   {
    $MeetingDuration="480"
   }
   else
   { 
    $MeetingDuration=($MeetingEndTime-$MeetingStartTime).TotalMinutes
   }
   $RoomUsage =$RoomUsage+$MeetingDuration
   $ReqiredAttendees=(($_.Attendees | Where {$_.Type -eq "required"}).emailaddress | select -ExpandProperty Address) -join ","
   $OptionalAttendees=(($_.Attendees | Where {$_.Type -eq "optional"}).emailaddress | select -ExpandProperty Address) -join ","
   $AllAttendeesCount=(($_.Attendees | Where {$_.Type -ne "resource"}).emailaddress | Measure-Object).Count

   #Filter for retrieving online meetings
   if(($OnlineMeetingOnly.IsPresent) -and ($IsOnlineMeeting -eq $false))
   {
    $Print=0
   }
   #View meetings from a specific organizer
   if(($OrgEmailId -ne "") -and ($OrgEmailId -ne $Organizer))
   {
    $Print=0
   }
   #Show Todays meetings only
   if(($ShowTodaysMeetingsOnly.IsPresent) -and ($MeetingStartTime -lt (Get-Date).Date))
   {
    $Print=0
   }

   #Detailed Report
   if($Print -eq 1)
   {
    $PrintedMeetings++
    $ExportResult=[PSCustomObject]@{'Room Name'=$RoomName;'Organizer'=$Organizer;'Subject'=$MeetingSubject;'Start Time'=$MeetingStartTime;'End Time'=$MeetingEndTime;'Duration(in mins)'=$MeetingDuration;'TimeZone'=$MeetingStartTimeZone;'Total Attendees Count'=$AllAttendeesCount;'Required Attendees'=$ReqiredAttendees;'Optional Attendees'=$OptionalAttendees;'Is Online Meeting'=$IsOnlineMeeting;'Is AllDay Meeting'=$IsAllDayMeeting}
    $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
   }
  }
 }  
 #Summary Report
    $ExportSummary=[PSCustomObject]@{'Room Name'=$RoomName;'Total Meeting Count'=$MeetingCount;'Online Meeting Count'=$OnlineMeetingCount;'Usage Duration(in mins)'=$RoomUsage;'Full Day Meetings'=$AllDayMeetingCount}
    $ExportSummary | Export-Csv -Path $ExportSummaryCSV -Notype -Append
}

#Open output file after execution
Write-Host `nScript executed successfully
if((Test-Path -Path $ExportCSV) -eq "True")
{
    Write-Host "Exported report has $PrintedMeetings meeting(s)" 
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6)
    {
        Invoke-Item "$ExportCSV"
        Invoke-Item "$ExportSummaryCSV"
    }
    Write-Host "Detailed report available in: $ExportCSV"
    Write-Host "Summary report available in: $ExportSummaryCSV"
}
else
{
    Write-Host "No meetings found" -ForegroundColor Red
}


    

   
  


 