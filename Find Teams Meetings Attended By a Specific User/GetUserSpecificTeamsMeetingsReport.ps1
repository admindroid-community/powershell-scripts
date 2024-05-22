<#
=============================================================================================

Name         : Find the Teams Meetings Attended by a Specific User Using PowerShell 
Version      : 1.0
website      : o365reports.com

-----------------
Script Highlights
-----------------
1. The script exports user specific Teams meeting report.  
2. The script retrieves user specific Teams meeting data for 180 days, by default. 
3. Allows you to obtain audit Teams meeting attended by specific user for a custom period. 
4. The script can be executed with an MFA enabled account too. 
5. It exports audit results to CSV file format in the working directory. 
6. Automatically installs the Exchange Online module (if not installed already) upon your confirmation. 
7. The script is scheduler friendly. 
8. The script supports certificate-based authentication.

For detailed Script execution:  https://o365reports.com/2024/05/22/find-the-teams-meetings-attended-by-a-specific-user-using-powershell/
============================================================================================
\#>

param (
    [Parameter(Mandatory = $false)]
    [string]$userUPN,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserName, 
    [string]$Password, 
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$Organization
)

Function Connect_Module {

    # Checking availability of Exchange Online module and installing module
    $ExchangeModule = Get-Module ExchangeOnlineManagement -ListAvailable
    if($ExchangeModule.count -eq 0)
    {
        Write-Host ExchangeOnline module is not available -ForegroundColor Yellow
        $confirm = Read-Host Do you want to Install ExchangeOnline module? [Y] Yes  [N] No
        if($confirm -match "[Yy]")
        {
            Write-Host "Installing ExchangeOnline module ..."
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
            Import-Module ExchangeOnlineManagement
        }
        else
        {
            Write-Host "ExchangeOnline Module is required. To Install ExchangeOnline module use 'Install-Module ExchangeOnlineManagement' cmdlet."
            Exit
        }
    }

    #Connecting to Exchange Online
    Write-Host "`nConnecting Exchange Online module ..." -ForegroundColor Yellow
    if(($UserName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
        Connect-ExchangeOnline -Credential $Credential 
    }
    elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "") {
        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
    }
    else {
        Connect-ExchangeOnline 
    }
}

#Verify user UPN is present
if($userUPN -eq "") {
    Write-Host User Principal Name is Required! -Foregroundcolor Yellow
    $userUPN = Read-Host Enter User UPN
}

$MaxStartDate=((Get-Date).AddDays(-180)).Date

#Getting Teams meeting attendance report for past 180 days
if(($StartDate -eq $null) -and ($EndDate -eq $null)) {
    $EndDate=(Get-Date)  #.Date
    $StartDate=$MaxStartDate
}

#Getting start date to generate Teams meetings attendance report
While($true) {
    if ($StartDate -eq $null) {
        $StartDate = Read-Host Enter start time for report generation '(Eg:MM/DD/YYYY)'
    }
    Try {
        $Date=[DateTime]$StartDate
        if($Date -ge $MaxStartDate) { 
            break
        }
        else {
            Write-Host `nAudit can be retrieved only for past 180 days. Please select a date after $MaxStartDate -ForegroundColor Red
            return
        }
    }
    Catch {
        Write-Host `nNot a valid date -ForegroundColor Red
    }
}

#Getting end date for teams attendance report
While($true) {
    if ($EndDate -eq $null) {
        $EndDate = Read-Host Enter End time for report generation '(Eg: MM/DD/YYYY)'
    }
    Try {
        $Date=[DateTime]$EndDate
        if($EndDate -lt ($StartDate))
        {
        Write-Host End time should be later than start time -ForegroundColor Red
        return
        }
        break
    }
    Catch {
        Write-Host `nNot a valid date -ForegroundColor Red
    }
}

#get file dir
$outputFilePath = $PSScriptRoot
$OutputCSV = "$outputFilePath\TeamsMeetingSpecificUserAttendanceReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$IntervalTimeInMinutes = 1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart = $StartDate
$CurrentEnd = $CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Check whether CurrentEnd exceeds EndDate
if($CurrentEnd -gt $EndDate) {
    $CurrentEnd = $EndDate
}

if($CurrentStart -eq $CurrentEnd) {
    Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
    Exit
}

Connect_Module

$CurrentResultCount = 0
$ResultCount = 0
Write-Host `nGenerating Teams meeting attendance report from $StartDate to $EndDate...
$ProcessedAuditCount = 0
$OutputEvents = 0
$ExportResult = ""   
$ExportResults = @()  
$RetriveOperation = "MeetingParticipantDetail"
while($true) { 
 #Getting Teams meeting participants report for the given time range
    Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $RetriveOperation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 | ForEach-Object {
        $ResultCount++
        $ProcessedAuditCount++
        Write-Progress -Activity "`n     Retrieving Team meeting attended by user : $userUPN ..."`n" Processed audit record count: $ProcessedAuditCount"
        $AuditData=$_.AuditData | ConvertFrom-Json
        $MeetingID=$AuditData.MeetingDetailId
        $CreatedBy=$_.UserIDs
        $AttendeesInfo=($AuditData.Attendees)
        $AttendeesType=$AttendeesInfo.RecipientType
        $Attendees=$AttendeesInfo.DisplayName
        $AttendeesUPN=$AttendeesInfo.UPN
        $JoinTime=(Get-Date($AuditData.JoinTime)).ToLocalTime()  #Get-Date($AuditData.JoinTime) Uncomment to view the Activity Time in UTC
        $LeaveTime=(Get-Date($AuditData.LeaveTime)).ToLocalTime()
        $Duration = $JoinTime - $LeaveTime
        $DurationinSeconds = ($Duration).TotalSeconds
        $TimeSpanDuration =  [timespan]::fromseconds($DurationinSeconds)
        $AttendedDuration = ("{0:hh\:mm\:ss}" -f $TimeSpanDuration)

        if($AttendeesUPN -eq $userUPN) {  
            #Export result to csv
            $OutputEvents++
            $ExportResult=@{'Meeting id'=$MeetingID;'Created By'=$CreatedBy;'Attendee'=$Attendees;'Attendee UPN' = $AttendeesUPN;'Attendee Type'=$AttendeesType;'Joined Time'=$JoinTime;'Left Time'=$LeaveTime;'Duration' = $AttendedDuration}
            $ExportResults= New-Object PSObject -Property $ExportResult
            $ExportResults | Select-Object 'Meeting id','Created By','Attendee','Attendee UPN','Attendee Type','Joined Time','Left Time','Duration' | Export-Csv -Path $OutputCSV -NoTypeInformation -Append 
        }
    }

    $currentResultCount=$currentResultCount+$ResultCount

    if($CurrentResultCount -ge 50000) {
        Write-Host Retrieved max record for current range.Proceeding further may cause data loss or rerun the script with reduced time interval. -ForegroundColor Red
        $Confirm=Read-Host `nAre you sure you want to continue? [Y] Yes [N] No
        if($Confirm -match "[Y]") {
            Write-Host Proceeding audit log collection with data loss
            [DateTime]$CurrentStart=$CurrentEnd
            [DateTime]$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
            $CurrentResultCount=0
            if($CurrentEnd -gt $EndDate) {
                $CurrentEnd=$EndDate
            }
        }
        else {
            Write-Host Please rerun the script with reduced time interval -ForegroundColor Red
            Exit
        }
    }

    if($ResultCount -lt 5000) { 
        if($CurrentEnd -eq $EndDate) {
            break
        }
        $CurrentStart=$CurrentEnd 
        if($CurrentStart -gt (Get-Date)) {
            break
        }
        $CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
        $CurrentResultCount=0
        if($CurrentEnd -gt $EndDate) {
            $CurrentEnd=$EndDate
        }
    }               

    $ResultCount=0
}

#Open output file after execution
if($OutputEvents -eq 0) {
    Write-Host No records found
}
else {
    Write-Host `nThe Teams meeting attendance report contains $OutputEvents audit records -ForegroundColor Green
    if((Test-Path -Path $OutputCSV) -eq "True")  {
        Write-Host `n "The output file available in: " -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCSV"
        Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
        Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
        Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
        $Prompt = New-Object -ComObject wscript.shell   
        $UserInput = $Prompt.popup("Do you want to open output file?",0,"Open Output File",4)   
        If ($UserInput -eq 6) {   
            Invoke-Item "$OutputCSV"   
        } 
    }
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue