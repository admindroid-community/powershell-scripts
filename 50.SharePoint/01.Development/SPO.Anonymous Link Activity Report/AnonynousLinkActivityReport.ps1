<#
=============================================================================================
Name:           SharePoint Online Anonymous Link Activity Report
Description:    This script exports SharePoint Online anonymous link activities report to CSV
Version:        3.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. Allow to generate 8 different anonymous link reports.
2. The script uses Exchange Online V3 module with REST API for secure connections without WinRM Basic Auth.
3. The script can be executed with MFA enabled account too.
4. Supports certificate-based authentication (CBA) for unattended scenarios.
5. Exports report results to CSV file.
6. Automatically installs the EXO V3 module (if not installed already) upon your confirmation.
7. The script is scheduler-friendly with multiple authentication options.
8. Improved error handling and progress reporting.
9. LoadCmdletHelp parameter support for accessing Get-Help cmdlet.
10. Enhanced output formatting and visual feedback.

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
    # Note: Password parameter kept as string for backward compatibility and automation scenarios
    # For better security, use certificate-based authentication instead
    [string]$Password,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$TenantId,
    [Switch]$LoadCmdletHelp,
    [Switch]$Help
)


# Show help if requested
if ($Help) {
    Write-Host @"
SYNOPSIS
    SharePoint Online Anonymous Link Activity Report

DESCRIPTION
    This script exports SharePoint Online anonymous link activities report to CSV using Exchange Online V3 module.
    The V3 module uses REST API for secure connections without requiring WinRM Basic Auth.

PARAMETERS
    -StartDate           : Start date for report (max 90 days back)
    -EndDate             : End date for report (default: today)
    -SharePointOnline    : Only SharePoint events
    -OneDrive            : Only OneDrive events
    -AnonymousSharing    : Only 'AnonymousLinkCreated' events
    -AnonymousAccess     : Only 'AnonymousLinkUsed' events
    -AdminName           : Username for basic authentication
    -Password            : Password for basic authentication
    -ClientId            : Client ID for app-based authentication
    -CertificateThumbprint : Certificate thumbprint for certificate-based authentication
    -TenantId            : Tenant ID for certificate-based authentication
    -LoadCmdletHelp      : Load Get-Help cmdlet functionality
    -Help                : Show this help message

AUTHENTICATION OPTIONS
    1. Interactive (Modern Auth) - Default method
    2. Basic Authentication - Using AdminName and Password
    3. Certificate-based Authentication - Using ClientId, CertificateThumbprint, and TenantId

EXAMPLES
    # Interactive authentication (recommended)
    .\AnonynousLinkActivityReport.ps1

    # SharePoint Online events only with date range
    .\AnonynousLinkActivityReport.ps1 -SharePointOnline -StartDate "2024-01-01" -EndDate "2024-01-31"

    # OneDrive events only
    .\AnonynousLinkActivityReport.ps1 -OneDrive

    # Anonymous sharing events only
    .\AnonynousLinkActivityReport.ps1 -AnonymousSharing

    # Anonymous access events only
    .\AnonynousLinkActivityReport.ps1 -AnonymousAccess

    # Certificate-based authentication
    .\AnonynousLinkActivityReport.ps1 -ClientId "app-id" -CertificateThumbprint "cert-thumbprint" -TenantId "tenant-id"

    # Load cmdlet help functionality
    .\AnonynousLinkActivityReport.ps1 -LoadCmdletHelp

NOTES
    - Audit data is available for the past 90 days only
    - The script uses Exchange Online V3 module with REST API
    - Certificate-based authentication is recommended for automation scenarios
    - Interactive authentication supports MFA
    - Can generate 8 different types of anonymous link reports based on filters

"@ -ForegroundColor Cyan
    exit 0
}

# Check for Exchange Online V3 module installation
$Module = Get-Module ExchangeOnlineManagement -ListAvailable
if($Module.count -eq 0)
{
    Write-Host "Exchange Online PowerShell V3 module is not available" -ForegroundColor Yellow
    $Confirm = Read-Host "Are you sure you want to install the module? [Y] Yes [N] No"
    if($Confirm -match "[yY]")
    {
        Write-Host "Installing Exchange Online PowerShell V3 module..." -ForegroundColor Magenta
        Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
        Import-Module ExchangeOnlineManagement -Force
    }
    else
    {
        Write-Host "EXO V3 module is required to connect to Exchange Online. Please install the module using: Install-Module ExchangeOnlineManagement" -ForegroundColor Red
        exit 1
    }
}

function Connect-ExchangeOnlineModern {
    try {
        Write-Host "Connecting to Exchange Online using V3 module..." -ForegroundColor Cyan

        # Certificate-based authentication (recommended for automation)
        if(($ClientId -ne "") -and ($CertificateThumbprint -ne "") -and ($TenantId -ne ""))
        {
            Write-Host "Using certificate-based authentication..." -ForegroundColor Yellow
            if($LoadCmdletHelp) {
                Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $TenantId -LoadCmdletHelp
            } else {
                Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $TenantId
            }
        }
        # Basic authentication with username/password (legacy)
        elseif(($AdminName -ne "") -and ($Password -ne ""))
        {
            Write-Host "Using basic authentication..." -ForegroundColor Yellow
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential = New-Object System.Management.Automation.PSCredential $AdminName, $SecuredPassword
            if($LoadCmdletHelp) {
                Connect-ExchangeOnline -Credential $Credential -LoadCmdletHelp
            } else {
                Connect-ExchangeOnline -Credential $Credential
            }
        }
        # Interactive authentication (default - supports MFA)
        else
        {
            Write-Host "Using interactive authentication (supports MFA)..." -ForegroundColor Yellow
            if($LoadCmdletHelp) {
                Connect-ExchangeOnline -LoadCmdletHelp
            } else {
                Connect-ExchangeOnline
            }
        }

        # Verify connection
        $orgConfig = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "✓ Successfully connected to Exchange Online: $($orgConfig.DisplayName)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}


# Date handling
$MaxStartDate = ((Get-Date).AddDays(-89)).Date
if(($null -eq $StartDate) -and ($null -eq $EndDate)) {
    $EndDate = (Get-Date).Date
    $StartDate = $MaxStartDate
}

while($true) {
    if ($null -eq $StartDate) {
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
        $StartDate = $null
    }
}

while($true) {
    if ($null -eq $EndDate) {
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
        $EndDate = $null
    }
}


$OutputCSV = "./AnonymousLinksActivityReport_" + ((Get-Date -format "yyyy-MMM-dd-ddd hh-mm tt").ToString()) + ".csv"
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
$ExportResults = @()

if($AnonymousSharing.IsPresent) {
    $RetriveOperation = "AnonymousLinkCreated"
    Write-Host "Filtering for Anonymous Sharing events only..." -ForegroundColor Yellow
} elseif($AnonymousAccess.IsPresent) {
    $RetriveOperation = "AnonymousLinkUsed"
    Write-Host "Filtering for Anonymous Access events only..." -ForegroundColor Yellow
} else {
    $RetriveOperation = "AnonymousLinkRemoved,AnonymousLinkCreated,AnonymousLinkUpdated,AnonymousLinkUsed"
    Write-Host "Retrieving all anonymous link activities..." -ForegroundColor Yellow
}

$CurrentResultCount = 0
while($true) {
    Write-Host "Processing time range: $CurrentStart to $CurrentEnd" -ForegroundColor Yellow
    $ResultCount = 0
    
    # Get anonymous sharing/access audit data for given time range
    try {
        $Results = Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $RetriveOperation -SessionId "AnonymousLinkAudit" -SessionCommand ReturnLargeSet -ResultSize 5000
    } catch {
        Write-Host "Error retrieving audit log: $($_.Exception.Message)" -ForegroundColor Red
        break
    }
    
    $ResultCount = ($Results | Measure-Object).Count
    foreach($Result in $Results) {
        $ProcessedAuditCount++
        $MoreInfo = $Result.AuditData
        $Operation = $Result.Operations
        $AuditData = $Result.AuditData | ConvertFrom-Json
        $Workload = $AuditData.Workload

        # Filter for SharePointOnline/OneDrive events
        if($SharePointOnline.IsPresent -and ($Workload -eq "OneDrive")) { 
            continue 
        }
        if($OneDrive.IsPresent -and ($Workload -eq "SharePoint")) { 
            continue 
        }

        $ActivityTime = (Get-Date($AuditData.CreationTime)).ToLocalTime()
        $PerformedBy = $AuditData.UserId
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
        $ExportResult = [PSCustomObject]@{
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
        $ExportResults += $ExportResult
    }
    Write-Progress -Activity "Retrieving anonymous link activities from $CurrentStart to $CurrentEnd" -Status "Processed audit record count: $ProcessedAuditCount" -PercentComplete (($ProcessedAuditCount % 100))
    $CurrentResultCount += $ResultCount
    
    if($CurrentResultCount -ge 50000) {
        Write-Host "Retrieved max record for current range. Proceeding further may cause data loss or rerun the script with reduced time interval." -ForegroundColor Red
        $Confirm = Read-Host "Are you sure you want to continue? [Y] Yes [N] No"
        if($Confirm -match "[yY]") {
            Write-Host "Proceeding audit log collection with potential data loss" -ForegroundColor Yellow
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


# Export all results to CSV
if($ExportResults.Count -gt 0) {
    $ExportResults | Export-Csv -Path $OutputCSV -NoTypeInformation
}

# Clear progress bar
Write-Progress -Activity "Completed" -Completed

if($OutputEvents -eq 0) {
    Write-Host "No records found" -ForegroundColor Yellow
} else {
    Write-Host "The output file contains $OutputEvents audit records" -ForegroundColor Green
    if((Test-Path -Path $OutputCSV) -eq $true) {
        Write-Host "Output file available in: " -NoNewline -ForegroundColor Yellow
        Write-Host $OutputCSV -ForegroundColor Cyan
        Write-Host "`n~~ Script prepared by AdminDroid Community ~~" -ForegroundColor Green
        Write-Host "~~ Check out " -NoNewline -ForegroundColor Green
        Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline
        Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green
        
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)
        If ($UserInput -eq 6) {
            Invoke-Item $OutputCSV
        }
    }
}

# Disconnect Exchange Online session
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
try {
    Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
    Write-Host "✓ Disconnected successfully" -ForegroundColor Green
} catch {
    # Ignore disconnect errors
}