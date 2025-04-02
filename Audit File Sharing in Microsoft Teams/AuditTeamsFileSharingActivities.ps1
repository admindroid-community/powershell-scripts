<#
=============================================================================================
Name:           Audit File Sharing Activities in Microsoft Teams Channels 
Version:        1.0
Website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the Exchange Online PowerShell module (if not already installed) upon your confirmation.
2. Tracks all file sharing activities in Microsoft Teams Channels across your tenant.
3. Finds files shared by external users in Microsoft Teams.
4. The script can be executed with an MFA-enabled account as well.
5. Finds files shared in a specific team and channel.
6. It exports the report result to CSV format.
7. It can be executed with Certificate-based Authentication too.
8. The script is scheduler friendly.


For detailed Script execution:  https://o365reports.com/2025/04/02/how-to-audit-file-sharing-activities-in-microsoft-teams/
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [switch]$ByExternalUsersOnly,
    [string]$TeamName,
    [string]$ChannelName,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password
)


$MaxStartDate=((Get-Date).AddDays(-180)).Date
#Retrive audit log for the past 180 days
if(($StartDate -eq $null) -and ($EndDate -eq $null)) {
    $EndDate=(Get-Date).Date
    $StartDate=$MaxStartDate
}
#Getting start date to audit export report
While($true) {
    if ($StartDate -eq $null) {
        $StartDate=Read-Host Enter start time for report generation '(Eg:09/24/2024)'
    }
    try {
        $Date=[DateTime]$StartDate
        if($Date -ge $MaxStartDate) { 
            break
        }
        else {
            Write-Host `nAudit can be retrieved only for the past 180 days. Please select a date after $($MaxStartDate.ToString("MM/dd/yyyy.")) -ForegroundColor Red
            return
        }
    }
    Catch {
        Write-Host `nNot a valid date -ForegroundColor Red
    }
}


#Getting end date to export audit report
While($true) {
    if ($EndDate -eq $null) {
        $EndDate=Read-Host Enter End time for report generation '(Eg: 9/24/2024)'
    }
    try {
        $Date=[DateTime]$EndDate
        if($EndDate -lt ($StartDate)) {
            Write-Host End time should be later than start time -ForegroundColor Red
            return
        }
        break
    }
    Catch {
        Write-Host `nNot a valid date -ForegroundColor Red
    }
}


Function Connect_Exo
{
    #Check for EXO module inatallation
    $Module = Get-Module ExchangeOnlineManagement -ListAvailable
    if($Module.count -eq 0)  { 
        Write-Host Exchange Online PowerShell  module is not available  -ForegroundColor yellow  
        $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
        if($Confirm -match "[yY]")  { 
            Write-host "Installing Exchange Online PowerShell module"
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
        } 
        else  { 
            Write-Host EXO module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
            Exit
        }
    } 
    Write-Host Connecting to Exchange Online...
    #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
    if(($UserName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
        Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
    }
    elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "") {
        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
}
Connect_Exo

$Location=Get-Location
$OutputCSV="$Location\Audit_Teams_File_Sharing_Activities_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Check whether CurrentEnd exceeds EndDate(checks for 1st iteration)
if($CurrentEnd -gt $EndDate) {
    $CurrentEnd=$EndDate
}

$AggregateResults = 0
$CurrentResult= @()
$CurrentResultCount=0
$FileSharingActivitiesCount=0
Write-Host `nRetrieving audit log from $StartDate to $EndDate...  -ForegroundColor Cyan

while($true) {
    if($CurrentStart -eq $CurrentEnd) {
        Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
        Exit
    }
    else { 
        $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations 'FileUploaded' -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
    }
 
    $AllAuditData=@()
    $AllAudits=$null
 
    foreach($Result in $Results) {
        $AuditData=$Result.auditdata | ConvertFrom-Json

        if (($AuditData.ApplicationDisplayName -match "Teams") -and ($AuditData.Workload -eq "SharePoint")) {
            
            $SharedTime = (Get-Date($AuditData.CreationTime)).ToLocalTime()
            $Team = $AuditData.SiteUrl -replace '^https://[^/]+/sites/([^/]+).*', '$1';
            $Channel = $AuditData.SourceRelativeUrl -replace '^Shared Documents/([^/]+).*', '$1'
            $SharedFile = $AuditData.SourceFileName
            $SharedBy = $AuditData.UserId
            $SharedFileURL = $AuditData.ObjectId
            $FileExtension = $AuditData.SourceFileExtension
            $SiteURL = $AuditData.SiteUrl
            $SharedFileRelativeURL = $AuditData.SourceRelativeUrl
            
            if (($ByExternalUsersOnly.IsPresent) -and ($AuditData.UserId -notmatch "#EXT#")) { continue }

            if ((!([string]::IsNullOrEmpty($TeamName))) -and ($TeamName -ne $Team)) { continue }

            if (!([string]::IsNullOrEmpty($ChannelName))) { 
                if (!([string]::IsNullOrEmpty($TeamName))) {
                    if (($TeamName -ne $Team) -or ($ChannelName -ne $Channel)) { 
                        continue 
                     }
                }
                else {
                    Write-Host "`nError: TeamName param is mandatory to filter based on Channels." -ForegroundColor Red
                    Exit
                }
            }

            $FileSharingActivitiesCount++
            $AllAudits=@{'Shared Time'=$SharedTime;'Shared File'=$SharedFile;'Shared by'=$SharedBy;'Site URL'=$SiteURL;'Team Name'=$Team;'Channel Name'=$Channel;'File Extension'=$FileExtension;'Shared File URL'=$SharedFileURL;'More Info'=$AuditData}
            $AllAuditData= New-Object PSObject -Property $AllAudits
            $AllAuditData | Sort 'Shared Time' | select 'Shared Time','Team Name','Channel Name','Shared File','Shared by','File Extension','Shared File URL','Site URL','More Info' | Export-Csv $OutputCSV -NoTypeInformation -Append
        }
    }
 
    $currentResultCount=$CurrentResultCount+($Results.count)
    $AggregateResults +=$Results.count
    Write-Progress -Activity "`n     Retrieving audit log for $CurrentStart : $CurrentResultCount records"`n" Total processed audit record count: $AggregateResults"
    if(($CurrentResultCount -eq 50000) -or ($Results.count -lt 5000)) {
        if($CurrentResultCount -eq 50000) {
            Write-Host Retrieved max record for the current range.Proceeding further may cause data loss or rerun the script with reduced time interval. -ForegroundColor Red
            $Confirm=Read-Host `nAre you sure you want to continue? [Y] Yes [N] No
            if($Confirm -notmatch "[Y]") {
                Write-Host Please rerun the script with reduced time interval -ForegroundColor Red
                Exit
            }
            else {
                Write-Host Proceeding audit log collection with data loss
            }
        }
        #Check for last iteration
        if(($CurrentEnd -eq $EndDate)) {
            break
        }
        [DateTime]$CurrentStart=$CurrentEnd
        #Break loop if start date exceeds current date(There will be no data)
        if($CurrentStart -gt (Get-Date)) {
            break
        }
        [DateTime]$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
        if($CurrentEnd -gt $EndDate) {
            $CurrentEnd=$EndDate
        }

        $CurrentResultCount=0
        $CurrentResult = @()
    }
}


Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green 
  

If($FileSharingActivitiesCount -eq 0) {
    Write-Host No records found
}
else {
    Write-Host `nThe output file contains $FileSharingActivitiesCount audit records.
    if((Test-Path -Path $OutputCSV) -eq "True") {
        Write-Host `n"The Output file available in: " -NoNewline -ForegroundColor Yellow
        Write-Host $OutputCSV 
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open output file?",0,"Open Output File",4)
        If ($UserInput -eq 6) {
            Invoke-Item "$OutputCSV"
        }
    }
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false