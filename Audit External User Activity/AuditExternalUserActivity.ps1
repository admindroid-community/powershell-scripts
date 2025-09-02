<#
=============================================================================================
Name:           Audit Office 365 external user activities report
Description:    This script exports external user activities report to CSV
Version:        3.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~~

1. The script uses Exchange Online V3 module with REST API for secure connections without WinRM Basic Auth.
2. The script can be executed with MFA enabled account too.   
3. Supports certificate-based authentication (CBA) for unattended scenarios.
4. Exports report results to CSV file.   
5. The script tracks all external users or a specific user activity based on the input. 
6. Allows you to generate an activity report for a custom period.   
7. Automatically installs the EXO V3 module (if not installed already) upon your confirmation.  
8. The script is scheduler-friendly with multiple authentication options.
9. Improved error handling and progress reporting.
10. LoadCmdletHelp parameter support for accessing Get-Help cmdlet.

For detailed script execution: https://o365reports.com/2022/02/10/audit-office-365-external-user-activities-using-powershell
============================================================================================
#>

param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$ExternalUserId,
    [string]$UserName,
    # Note: Password parameter kept as string for backward compatibility and automation scenarios
    # For better security, use certificate-based authentication instead
    [string]$Password,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$TenantId,
    [Switch]$LoadCmdletHelp,
    [Switch]$Help
)

if ($Help) {
    Write-Host @"
SYNOPSIS
    Audit Office 365 External User Activities Report

DESCRIPTION
    This script exports external user activities report to CSV using Exchange Online V3 module.
    The V3 module uses REST API for secure connections without requiring WinRM Basic Auth.

PARAMETERS
    -StartDate          : Start date for audit report (max 90 days back)
    -EndDate            : End date for audit report
    -ExternalUserId     : Filter by specific external user email
    -UserName           : Username for basic authentication
    -Password           : Password for basic authentication
    -ClientId           : Client ID for app-based authentication
    -CertificateThumbprint : Certificate thumbprint for certificate-based authentication
    -TenantId           : Tenant ID for certificate-based authentication
    -LoadCmdletHelp     : Load Get-Help cmdlet functionality
    -Help               : Show this help message

AUTHENTICATION OPTIONS
    1. Interactive (Modern Auth) - Default method
    2. Basic Authentication - Using UserName and Password
    3. Certificate-based Authentication - Using ClientId, CertificateThumbprint, and TenantId

EXAMPLES
    # Interactive authentication (recommended)
    .\AuditExternalUserActivity.ps1

    # With date range
    .\AuditExternalUserActivity.ps1 -StartDate "2024-01-01" -EndDate "2024-01-31"

    # Filter by specific external user
    .\AuditExternalUserActivity.ps1 -ExternalUserId "user@external.com"

    # Certificate-based authentication
    .\AuditExternalUserActivity.ps1 -ClientId "app-id" -CertificateThumbprint "cert-thumbprint" -TenantId "tenant-id"

    # Load cmdlet help functionality
    .\AuditExternalUserActivity.ps1 -LoadCmdletHelp

NOTES
    - Audit data is available for the past 90 days only
    - The script uses Exchange Online V3 module with REST API
    - Certificate-based authentication is recommended for automation scenarios
    - Interactive authentication supports MFA

"@ -ForegroundColor Cyan
    exit 0
}

Function Connect_Exo
{
    # Check for EXO V3 module installation
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
            Exit
        }
    }

    Write-Host "Connecting to Exchange Online using V3 module..." -ForegroundColor Cyan

    try {
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
        elseif(($UserName -ne "") -and ($Password -ne ""))
        {
            Write-Host "Using basic authentication..." -ForegroundColor Yellow
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword
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

$MaxStartDate = ((Get-Date).AddDays(-89)).Date

# Getting external user activity for past 90 days
if(($null -eq $StartDate) -and ($null -eq $EndDate))
{
    $EndDate = (Get-Date).Date
    $StartDate = $MaxStartDate
}

# Getting start date to generate external user activity report
While($true)
{
    if ($null -eq $StartDate)
    {
        $StartDate = Read-Host "Enter start time for report generation (Eg:04/28/2021)"
    }
    Try
    {
        $Date = [DateTime]$StartDate
        if($Date -ge $MaxStartDate)
        { 
            break
        }
        else
        {
            Write-Host "`nAudit can be retrieved only for past 90 days. Please select a date after $MaxStartDate" -ForegroundColor Red
            return
        }
    }
    Catch
    {
        Write-Host "`nNot a valid date" -ForegroundColor Red
        $StartDate = $null
    }
}

# Getting end date to audit external user activity report
While($true)
{
    if ($null -eq $EndDate)
    {
        $EndDate = Read-Host "Enter End time for report generation (Eg: 04/28/2021)"
    }
    Try
    {
        $Date = [DateTime]$EndDate
        if($EndDate -lt $StartDate)
        {
            Write-Host "End time should be later than start time" -ForegroundColor Red
            return
        }
        break
    }
    Catch
    {
        Write-Host "`nNot a valid date" -ForegroundColor Red
        $EndDate = $null
    }
}

$OutputCSV = ".\AuditExternalUserActivityReport_$((Get-Date -format "yyyy-MMM-dd-ddd hh-mm tt").ToString()).csv" 
$IntervalTimeInMinutes = 1440    # 24 hours
$CurrentStart = $StartDate
$CurrentEnd = $CurrentStart.AddMinutes($IntervalTimeInMinutes)

# Check whether CurrentEnd exceeds EndDate
if($CurrentEnd -gt $EndDate)
{
    $CurrentEnd = $EndDate
}

if($CurrentStart -eq $CurrentEnd)
{
    Write-Host "Start and end time are same. Please enter different time range" -ForegroundColor Red
    Exit
}

# Connect to Exchange Online
$connectionResult = Connect_EXO
if (-not $connectionResult) {
    Write-Host "Failed to connect to Exchange Online. Exiting..." -ForegroundColor Red
    Exit
}

$CurrentResultCount = 0
$ProcessedAuditCount = 0
$OutputEvents = 0
$ExportResults = @()  

if($ExternalUserId -eq "")
{
    $UserId = "*#EXT*"
}
else
{
    $UserId = $ExternalUserId
}

Write-Host "`nAuditing external user activities from $StartDate to $EndDate..." -ForegroundColor Cyan

while($true)
{ 
    Write-Host "Processing time range: $CurrentStart to $CurrentEnd" -ForegroundColor Yellow
    $ResultCount = 0
    
    # Getting external user audit data for the given time range
    try {
        $auditLogs = Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -UserIds $UserId -SessionId "ExternalUserAudit" -SessionCommand ReturnLargeSet -ResultSize 5000
        
        foreach($log in $auditLogs) {
            $ResultCount++
            $ProcessedAuditCount++
            Write-Progress -Activity "Retrieving external user activities from $CurrentStart to $CurrentEnd" -Status "Processed audit record count: $ProcessedAuditCount" -PercentComplete (($ProcessedAuditCount % 100))
            
            $MoreInfo = $log.AuditData
            $AuditData = $log.AuditData | ConvertFrom-Json
            $ActivityTime = (Get-Date($AuditData.CreationTime)).ToLocalTime()  # Convert to local time
            $UserID = $AuditData.UserId
            $Operation = $AuditData.Operation
            $ResourceType = $AuditData.ItemType
            $ResourceURL = $AuditData.ObjectId
            $Workload = $AuditData.Workload
            
            # Export result to csv
            $OutputEvents++
            $ExportResult = [PSCustomObject]@{
                'Activity Time' = $ActivityTime
                'User Name' = $UserID
                'Operation' = $Operation
                'Resource URL' = $ResourceURL
                'Resource Type' = $ResourceType
                'Workload' = $Workload
                'More Info' = $MoreInfo
            }
            $ExportResults += $ExportResult
        }
    }
    catch {
        Write-Host "Error retrieving audit logs: $($_.Exception.Message)" -ForegroundColor Red
        break
    }
    
    $CurrentResultCount = $CurrentResultCount + $ResultCount
    
    if($CurrentResultCount -ge 50000)
    {
        Write-Host "Retrieved max record for current range. Proceeding further may cause data loss or rerun the script with reduced time interval." -ForegroundColor Red
        $Confirm = Read-Host "`nAre you sure you want to continue? [Y] Yes [N] No"
        if($Confirm -match "[Y]")
        {
            Write-Host "Proceeding audit log collection with potential data loss" -ForegroundColor Yellow
            [DateTime]$CurrentStart = $CurrentEnd
            [DateTime]$CurrentEnd = $CurrentStart.AddMinutes($IntervalTimeInMinutes)
            $CurrentResultCount = 0
            if($CurrentEnd -gt $EndDate)
            {
                $CurrentEnd = $EndDate
            }
        }
        else
        {
            Write-Host "Please rerun the script with reduced time interval" -ForegroundColor Red
            Exit
        }
    }

    if($ResultCount -lt 5000)
    { 
        if($CurrentEnd -eq $EndDate)
        {
            break
        }
        $CurrentStart = $CurrentEnd 
        if($CurrentStart -gt (Get-Date))
        {
            break
        }
        $CurrentEnd = $CurrentStart.AddMinutes($IntervalTimeInMinutes)
        $CurrentResultCount = 0
        if($CurrentEnd -gt $EndDate)
        {
            $CurrentEnd = $EndDate
        }
    }                                                                                             
    $ResultCount = 0
}

# Export all results to CSV
if($ExportResults.Count -gt 0) {
    $ExportResults | Export-Csv -Path $OutputCSV -NoTypeInformation
}

# Clear progress bar
Write-Progress -Activity "Completed" -Completed

# Open output file after execution
If($OutputEvents -eq 0)
{
    Write-Host "No records found" -ForegroundColor Yellow
}
else
{
    Write-Host "`nThe output file contains $OutputEvents audit records" -ForegroundColor Green
    if((Test-Path -Path $OutputCSV) -eq "True") 
    {
        Write-Host "The Output file available in: " -NoNewline -ForegroundColor Yellow
        Write-Host $OutputCSV -ForegroundColor Cyan
        Write-Host "`n~~ Script prepared by AdminDroid Community ~~" -ForegroundColor Green
        Write-Host "~~ Check out " -NoNewline -ForegroundColor Green
        Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline
        Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green
        
        $Prompt = New-Object -ComObject wscript.shell   
        $UserInput = $Prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)   
        If ($UserInput -eq 6)   
        {   
            Invoke-Item "$OutputCSV"   
        } 
    }
}

# Disconnect Exchange Online session
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
Write-Host "✓ Disconnected successfully" -ForegroundColor Green