<#
=============================================================================================

Name         : Export all Entra App Operations Using PowerShell  
Version      : 1.0
website      : o365reports.com

-----------------
Script Highlights
-----------------

1. Tracks all Entra app operations for the past 180 days.
2. Allows to track app operations for a custom date range.
3. The script automatically verifies and installs the Exchange Online PowerShell V3 module (if not installed already) upon your confirmation.
4. Enables filtering of app operations from the following categories.
    -> Added Applications
    -> Updated Applications
    -> Deleted Applications
    -> Consent to Applications
    -> OAuth2 Permission Grants
    -> App Role Assignments
    -> Service Principal Changes
    -> Credential Changes
    -> Delegation Changes
5. Tracks app operations performed on the specific application.
6. Generates a report that retrieves successful operations alone.
7. Helps export failed operations alone.
8. Audit app operations performed by a specific user.
10. The script can be executed with an MFA-enabled account too.
11. Exports report results as a CSV file.
12. The script is scheduler friendly.
13. It can be executed with certificate-based authentication (CBA) too.  

For detailed Script execution: https://o365reports.com/2025/05/27/monitor-entra-app-operations-using-powershell/
============================================================================================
\#>

Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [ValidateSet(
    "Added Applications","Updated Applications", "Deleted Applications",
    "Consent to Applications", "OAuth2 Permission Grant", "App Role Assignments",
    "Service Principal Changes", "Credential Changes", "Delegation Changes"
    )]
    [string[]]$Operations,
    [string]$AppId,
    [string]$PerformedBy,
    [switch]$SucceedOnly,
    [switch]$FailedOnly,
    [string]$UserName,
    [string]$Password,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbPrint
)

Function Connect_Module {
    # Check for Exchange Online module installation
    $ExchangeModule = Get-Module ExchangeOnlineManagement -ListAvailable
    if ($ExchangeModule.count -eq 0) {
        Write-Host "ExchangeOnline module is not available" -ForegroundColor Yellow
        $confirm = Read-Host "Do you want to Install ExchangeOnline module? [Y] Yes  [N] No"
        if ($confirm -match "[Yy]") {
            Write-Host "Installing ExchangeOnline module ..."
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
            Import-Module ExchangeOnlineManagement
        } else {
            Write-Host "ExchangeOnline Module is required. To Install ExchangeOnline module use 'Install-Module ExchangeOnlineManagement' cmdlet."
            Exit
        }
    }

    # Connect Exchange Online module
    Write-Host "`nConnecting Exchange Online module ..." -ForegroundColor Yellow
    if (($UserName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword
        Connect-ExchangeOnline -Credential $Credential
    } elseif ($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "") {
        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
    } else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
}

$MaxStartDate = ((Get-Date).AddDays(-180)).Date

# Default date range
if (($StartDate -eq $null) -and ($EndDate -eq $null)) {
    $EndDate = (Get-Date).Date
    $StartDate = $MaxStartDate
}

# Get start date
While ($true) {
    if ($StartDate -eq $null) {
        $StartDate = Read-Host "Enter start time for report generation '(Eg: 04/28/2021)'"
    }
    Try {
        $Date = [DateTime]$StartDate
        if ($Date -ge $MaxStartDate) {
            break
        } else {
            Write-Host "`nAudit can be retrieved only for past 180 days. Please select a date after $MaxStartDate" -ForegroundColor Red
            return
        }
    } Catch {
        Write-Host "`nNot a valid date" -ForegroundColor Red
    }
}

# Get end date
While ($true) {
    if ($EndDate -eq $null) {
        $EndDate = Read-Host "Enter End time for report generation '(Eg: 04/28/2021)'"
    }
    Try {
        $Date = [DateTime]$EndDate
        if ($EndDate -lt ($StartDate)) {
            Write-Host "End time should be later than start time" -ForegroundColor Red
            return
        }
        break
    } Catch {
        Write-Host "`nNot a valid date" -ForegroundColor Red
    }
}

$Location = Get-Location
$OutputCSV = "$Location\Audit_Application_Operation_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
$IntervalTimeInMinutes = 1440
$CurrentStart = $StartDate
$CurrentEnd = $CurrentStart.AddMinutes($IntervalTimeInMinutes)

if ($CurrentEnd -gt $EndDate) {
    $CurrentEnd = $EndDate
}

if ($CurrentStart -eq $CurrentEnd) {
    Write-Host "Start and end time are the same. Please enter a different time range." -ForegroundColor Red
    Exit
}

Connect_Module

# Initialize variables
$CurrentResultCount = 0
$AggregateResultCount = 0
Write-Host "`nAuditing application operations from $StartDate to $EndDate..."
$ExportResult = ""
$ExportResults = @()
$OutputEvents = 0
$OperationMap = @{
    "Added Applications"        = "Add application"
    "Updated Applications"      = "Update application"
    "Deleted Applications"      = "Delete application"
    "Consent to Applications"   = "Consent to application"
    "OAuth2 Permission Grant"   = "Add Oauth2PermissionGrant"
    "App Role Assignments"      = "Add app role assignment to service principal, Add app role assignment grant to user"
    "Service Principal Changes" = "Add service principal, Remove service principal, Update service principal"
    "Credential Changes"        = "Add service principal credentials, Remove service principal credentials"
    "Delegation Changes"        = "adddelegatedpermissiongrant, removedelegatedpermissiongrant"
}

if ($Operations){
    $OperationNames = $OperationMap[$Operations]
}
else{
    $OperationNames = "Add delegation entry, Add service principal, Add service principal credentials, Remove delegation entry, Remove service principal, Remove service principal credentials, Set delegation entry, Add application, Update application, Delete application, Consent to application, Add Oauth2PermissionGrant, Add app role assignment to service principal, Add app role assignment grant to user, Update service principal, Remove OAuth2PermissionGrant, addapproleassignmenttogroup, addownertoapplication, addownertoserviceprincipal, createapplicationpasswordforuser, deleteapplicationpasswordforuser, harddeleteapplication, removeapproleassignmentfromgroup, removeapproleassignmentfromserviceprincipal, removeapproleassignmentfromuser, revokeconsent, removeownerfromapplication, restoreapplication, removeownerfromserviceprincipal, adddelegatedpermissiongrant, removedelegatedpermissiongrant, Hard delete service principal, Update application – Certificates and secrets management"
}

# Main loop
while ($true) {
    $Results = Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $OperationNames -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
    $ResultCount = ($Results | Measure-Object).count

    ForEach ($Result in $Results) {
        $AggregateResultCount++
        $EventTime = $Result.CreationDate
        $ActivityBy = $Result.UserIds
        $Operation = $Result.Operations
        $MoreInfo = $Result.auditdata
        $AuditData = $Result.auditdata | ConvertFrom-Json
        $EventId = $AuditData.Id
        $ResultStatus = $AuditData.ResultStatus
        $TagetApp = $AuditData.Target[3].ID
        $TargetTypes = $AuditData.Target[2].ID
        $ModifiedProperties = if([string]::IsNullOrEmpty($AuditData.ModifiedProperties)) { '-' } else { $AuditData.ModifiedProperties | ConvertTo-Json -Compress }
        $AdditionalDetails = $AuditData.ExtendedProperties | Where-Object { $_.Name -eq 'additionalDetails' }
        $ApplicationId = ($AdditionalDetails.Value | ConvertFrom-Json).AppId
        #$IsAdminConsented = $AuditData.ModifiedProperties[0].NewValue | 'Is Admin Consented' = $IsAdminConsented; | 'Is Admin Consented', 

        # If no Type 5 found, use Type 1
        #$InitiatedBy = $AuditData.Actor | Where-Object { $_.Type -eq 5 } | Select-Object -ExpandProperty ID
        #if (-not $InitiatedBy) { $InitiatedBy = ($AuditData.Actor | Where-Object { $_.Type -eq 1 } | Select-Object -ExpandProperty ID) }

        $PrintFlag = "True"
        if(($PerformedBy.Length -ne 0) -and ($PerformedBy -ne $ActivityBy)) { $PrintFlag = "False" }
        if(($AppId.Length -ne 0) -and ($AppId -ne $ApplicationId)) { $PrintFlag = "False" }
        if($SucceedOnly.IsPresent -and ($ResultStatus -ne "Success")) { $PrintFlag = "False" }
        if($FailedOnly.IsPresent -and ($ResultStatus -ne "Failure")) { $PrintFlag = "False" }


        if($PrintFlag -eq "True")
        {
            $OutputEvents++
            $ExportResult = @{'Event Time' = $EventTime; 'Target Application' = $TagetApp; 'Target Type' = $TargetTypes; 'Operation' = $Operation; 'Performed By' = $ActivityBy; 'Result Status' = $ResultStatus; 'Modified Properties' = $ModifiedProperties; 'Event Id' = $EventId; 'More Info' = $MoreInfo}
            $ExportResults = New-Object PSObject -Property $ExportResult  
            $ExportResults | Sort 'Activity Time'| Select-Object 'Event Time', 'Operation','Target Application', 'Performed By', 'Target Type', 'Result Status', 'Modified Properties', 'Event Id', 'More Info' | Export-Csv -Path $OutputCSV -NoTypeInformation -Append 
        }
    }

    $currentResultCount=$CurrentResultCount+$ResultCount
    Write-Progress -Activity "`n     Retrieving application operation audit log for $CurrentStart : $CurrentResultCount records"`n" Processed audit record count: $AggregateResultCount"

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

Write-Host "`n~~ Script prepared by AdminDroid Community ~~" -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green

# Output results
if ($OutputEvents -eq 0) {
    Write-Host "`n No records found"
} else {
    Write-Host "`nThe output file contains $OutputEvents audit records"
    if ((Test-Path -Path $OutputCSV) -eq $true) {
        Write-Host "`nThe Output file is available at: " -NoNewline -ForegroundColor Yellow; Write-Host $OutputCSV
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)
        If ($UserInput -eq 6) {
            Invoke-Item "$OutputCSV"
        }
    }
}

# Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue