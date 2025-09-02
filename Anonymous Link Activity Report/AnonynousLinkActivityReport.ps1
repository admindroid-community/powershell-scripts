<#
=============================================================================================
Name:           SharePoint Online Anonymous Link Activity Report
Description:    This script exports SharePoint Online anonymous link  activities report to CSV
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.Allow to generate 8 different anonymous link reports. 
2.The script uses modern authentication to retrieve audit logs.   
3.The script can be executed with MFA enabled account too.   
4.Exports report results to CSV file.   
5.Automatically installs the EXO V2 module (if not installed already) upon your confirmation.  
6.The script is scheduler friendly. I.e., Credential can be passed as a parameter instead of saving inside the script. 

For detailed script execution: https://o365reports.com/2021/06/22/audit-anonymous-access-in-sharepoint-online-using-powershell
============================================================================================
#>

param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [switch]$SharePointOnline,
    [switch]$OneDrive,
    [switch]$AnonymousSharing,
    [switch]$AnonymousAccess,
    [string]$AdminName,
    [string]$Password,
    [Switch]$Help
)


# Show help if requested
if ($Help) {
    Write-Host @"
SYNOPSIS
    Export SharePoint Online Anonymous Link Activity Report to CSV

DESCRIPTION
    This script exports SharePoint Online anonymous link activities report to CSV. Supports modern authentication, MFA, and scheduler-friendly credential passing.

PARAMETERS
    -StartDate           : Start date for report (default: 90 days ago)
    -EndDate             : End date for report (default: today)
    -SharePointOnline    : Only SharePoint events
    -OneDrive            : Only OneDrive events
    -AnonymousSharing    : Only 'AnonymousLinkCreated' events
    -AnonymousAccess     : Only 'AnonymousLinkUsed' events
    -AdminName           : Username for authentication
    -Password            : Password for authentication
    -Help                : Show this help message

EXAMPLES
    .\AnonynousLinkActivityReport.ps1 -SharePointOnline -StartDate "2025-08-01" -EndDate "2025-08-31"
    .\AnonynousLinkActivityReport.ps1 -AnonymousSharing
    .\AnonynousLinkActivityReport.ps1 -AdminName "admin@tenant.com" -Password "yourpassword"
"@ -ForegroundColor Cyan
    exit 0
}

# Check for ExchangeOnlineManagement module
$Module = (Get-Module ExchangeOnlineManagement -ListAvailable).Name
if($null -eq $Module)
{
    Write-Host "Exchange Online PowerShell V2 module is not available." -ForegroundColor Yellow
    $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No: "
    if($Confirm -match "[yY]")
    {
        Write-Host "Installing Exchange Online PowerShell module..." -ForegroundColor Magenta
        Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
        Import-Module ExchangeOnlineManagement -Force
    }
    else
    {
        Write-Host "Exiting. Exchange Online PowerShell module must be available to run the script." -ForegroundColor Red
        exit 1
    }
}

function Connect-ExchangeOnlineModern {
    try {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        if(($AdminName -ne "") -and ($Password -ne "")) {
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential  = New-Object System.Management.Automation.PSCredential $AdminName,$SecuredPassword
            Connect-ExchangeOnline -Credential $Credential
        } else {
            Connect-ExchangeOnline
        }
        return $true
    } catch {
        Write-Host "Error connecting to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}


# Date handling
$MaxStartDate = ((Get-Date).AddDays(-89)).Date
if(($StartDate -eq $null) -and ($EndDate -eq $null)) {
    $EndDate = (Get-Date).Date
    $StartDate = $MaxStartDate
}

while($true) {
    if ($StartDate -eq $null) {
        $StartDate = Read-Host "Enter start time for report generation (Eg: 04/28/2021)"
    }
    try {
        $Date = [DateTime]$StartDate
        if($Date -ge $MaxStartDate) {
            break
        } else {
            Write-Host "Anonymous activity report can be retrieved only for past 90 days. Please select a date after $MaxStartDate" -ForegroundColor Red
            return
        }
    } catch {
        Write-Host "Not a valid date" -ForegroundColor Red
    }
}

while($true) {
    if ($EndDate -eq $null) {
        $EndDate = Read-Host "Enter End time for report generation (Eg: 04/28/2021)"
    }
    try {
        $Date = [DateTime]$EndDate
        if($EndDate -lt $StartDate) {
            Write-Host "End time should be later than start time" -ForegroundColor Red
            return
        }
        break
    } catch {
        Write-Host "Not a valid date" -ForegroundColor Red
    }
}


$OutputCSV = "./AnonymousLinksActivityReport_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
$IntervalTimeInMinutes = 1440
$CurrentStart = $StartDate
$CurrentEnd = $CurrentStart.AddMinutes($IntervalTimeInMinutes)
if($CurrentEnd -gt $EndDate) { $CurrentEnd = $EndDate }
if($CurrentStart -eq $CurrentEnd) {
    Write-Host "Start and end time are same. Please enter different time range" -ForegroundColor Red
    exit 1
}

# Connect to Exchange Online
$connected = Connect-ExchangeOnlineModern
if (-not $connected) {
    Write-Host "Failed to connect to Exchange Online. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host "Retrieving anonymous link events from $StartDate to $EndDate..." -ForegroundColor Cyan
$ProcessedAuditCount = 0
$OutputEvents = 0
$ExportResult = ""
$ExportResults = @()
if($AnonymousSharing.IsPresent) {
    $RetriveOperation = "AnonymousLinkCreated"
} elseif($AnonymousAccess.IsPresent) {
    $RetriveOperation = "AnonymousLinkUsed"
} else {
    $RetriveOperation = "AnonymousLinkRemoved,AnonymousLinkCreated,AnonymousLinkUpdated,AnonymousLinkUsed"
}


$CurrentResultCount = 0
while($true) {
    # Get anonymous sharing/access audit data for given time range
    try {
        $Results = Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $RetriveOperation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
    } catch {
        Write-Host "Error retrieving audit log: $($_.Exception.Message)" -ForegroundColor Red
        break
    }
    $ResultCount = ($Results | Measure-Object).Count
    foreach($Result in $Results) {
        $ProcessedAuditCount++
        $MoreInfo = $Result.auditdata
        $Operation = $Result.Operations
        $AuditData = $Result.auditdata | ConvertFrom-Json
        $Workload = $AuditData.Workload

        # Filter for SharePointOnline/OneDrive events
        if($SharePointOnline.IsPresent -and ($Workload -eq "OneDrive")) { continue }
        if($OneDrive.IsPresent -and ($Workload -eq "SharePoint")) { continue }

        $ActivityTime = Get-Date($AuditData.CreationTime) -format g
        $PerformedBy = $AuditData.userId
        $ResourceType = $AuditData.ItemType
        $Resource = $AuditData.ObjectId
        $SiteURL = $AuditData.SiteURL
        $UserIP = $AuditData.ClientIP
        $EventData = $AuditData.EventData

        # Check whether the anonymous link has edit permission
        if($Operation -ne "AnonymousLinkUsed") {
            if($EventData -like "*View*") {
                $EditEnabled = "False"
            } else {
                $EditEnabled = "True"
            }
        } else {
            $EditEnabled = "NA"
        }

        # Export result to csv
        $OutputEvents++
        $ExportResult = @{
            'Activity Time' = $ActivityTime
            'Activity' = $Operation
            'Performed By' = $PerformedBy
            'User IP' = $UserIP
            'Resource Type' = $ResourceType
            'Shared/Accessed Resource' = $Resource
            'Edit Enabled' = $EditEnabled
            'Site URL' = $SiteURL
            'Workload' = $Workload
            'More Info' = $MoreInfo
        }
        $ExportResults = New-Object PSObject -Property $ExportResult
        $ExportResults | Select-Object 'Activity Time','Activity','Performed By','User IP','Resource Type','Shared/Accessed Resource','Edit Enabled','Site URL','Workload','More Info' | Export-Csv -Path $OutputCSV -Append -NoTypeInformation
    }
    Write-Progress -Activity "Retrieving anonymous link activities from $CurrentStart to $CurrentEnd..." -Status "Processed audit record count: $ProcessedAuditCount"
    $CurrentResultCount += $ResultCount
    if($CurrentResultCount -ge 50000) {
        Write-Host "Retrieved max record for current range. Proceeding further may cause data loss or rerun the script with reduced time interval." -ForegroundColor Red
        $Confirm = Read-Host "Are you sure you want to continue? [Y] Yes [N] No: "
        if($Confirm -match "[yY]") {
            Write-Host "Proceeding audit log collection with data loss" -ForegroundColor Yellow
            $CurrentStart = $CurrentEnd
            $CurrentEnd = $CurrentStart.AddMinutes($IntervalTimeInMinutes)
            $CurrentResultCount = 0
            if($CurrentEnd -gt $EndDate) { $CurrentEnd = $EndDate }
        } else {
            Write-Host "Please rerun the script with reduced time interval" -ForegroundColor Red
            break
        }
    }
    if($Results.Count -lt 5000) {
        if($CurrentEnd -eq $EndDate) { break }
        $CurrentStart = $CurrentEnd
        if($CurrentStart -gt (Get-Date)) { break }
        $CurrentEnd = $CurrentStart.AddMinutes($IntervalTimeInMinutes)
        $CurrentResultCount = 0
        if($CurrentEnd -gt $EndDate) { $CurrentEnd = $EndDate }
    }
}


if($OutputEvents -eq 0) {
    Write-Host "No records found" -ForegroundColor Yellow
} else {
    Write-Host "The output file contains $OutputEvents audit records" -ForegroundColor Green
    if((Test-Path -Path $OutputCSV) -eq $true) {
        Write-Host "Output file available in: " -NoNewline -ForegroundColor Yellow
        Write-Host $OutputCSV
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)
        If ($UserInput -eq 6) {
            Invoke-Item $OutputCSV
        }
    }
}

#Disconnect Exchange Online session
try {
    Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
    Write-Host "Disconnected from Exchange Online" -ForegroundColor Cyan
} catch {
    # Ignore disconnect errors
}