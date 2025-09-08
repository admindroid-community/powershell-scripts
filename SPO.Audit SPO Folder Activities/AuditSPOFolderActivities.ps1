<#
=============================================================================================
Name:           Monitor Folder Activities in SharePoint Online Using PowerShell   
Version:        1.0
Website:        o365reports.com


Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Tracks the folder activities in SharePoint and OneDrive for the past 180 days. 
2. Allows to track folder activities for a custom date range. 
3. Filters folder activities of SharePoint and OneDrive separately. 
4. Audit folder activities for a single site in SharePoint Online.  
5. Monitors folder activities for list of sites in SharePoint Online.  
6. Helps to audit folder activities by a specific user.  
7. Excludes system activities by default, with an option to include them if required.  
8. Exports report result into CSV file.  
9. The script automatically verifies and installs the Exchange Online PowerShell module (if not installed already) upon your confirmation. 
10. The script can be executed with an MFA-enabled account too. 
11. The script supports Certificate-based authentication (CBA). 
12. The script is scheduler friendly. 

For detailed Script execution: https://o365reports.com/2024/12/03/monitor-folder-activities-in-sharepoint-online-using-powershell/


============================================================================================
#>Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$PerformedBy,
    [string]$SiteUrl,
    [string]$ImportSitesCsv,
    [switch]$SharePointOnline,
    [switch]$OneDrive,
    [string]$UserName,
    [string]$Password,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbPrint,
    [switch]$IncludeSystemEvent
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
$OutputCSV = "$Location\Audit_SPO_Folder_Activity_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
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
$FilterSites = @()
$CurrentResultCount = 0
$AggregateResultCount = 0
Write-Host "`nAuditing folder activities from $StartDate to $EndDate..."
$ExportResult = ""
$ExportResults = @()
$OutputEvents = 0
$OperationNames = "FolderCreated, FolderModified, FolderRenamed, FolderCopied, FolderMoved, FolderDeleted, FolderRecycled, FolderDeletedFirstStageRecycleBin, FolderDeletedSecondStageRecycleBin, FolderRestored"

if ($ImportSitesCsv.Length -ne 0)
{
 $FilterSites = Import-Csv -Path $ImportSitesCsv | Select-Object -ExpandProperty SiteUrl
}
# Main loop
while ($true) {
    $Results = Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $OperationNames -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
    $ResultCount = ($Results | Measure-Object).count

    ForEach ($Result in $Results) {
        $AggregateResultCount++
        $MoreInfo = $Result.auditdata
        $Operation = $Result.Operations
        $ActionBy = $Result.UserIds
        $AuditData = $Result.auditdata | ConvertFrom-Json
        $Workload = $AuditData.Workload
        $Site = $AuditData.SiteUrl
        $PrintFlag = "True"

        # Excluding system events
        if (-not $IncludeSystemEvent) {
            if ($ActionBy -in @("app@sharepoint", "SHAREPOINT\system"))
            {
             $PrintFlag = "False"
            }
        }
        
        # Audit folder activities by specific user
        if(($PerformedBy.Length -ne 0) -and ($PerformedBy -ne $ActionBy))
        {
         $PrintFlag = "False"
        }

        # Filter for workload based folder usage events
        if($SharePointOnline.IsPresent -and ($Workload -eq "OneDrive"))
        {
         $PrintFlag = "False"
        }

        if($OneDrive.IsPresent -and ($Workload -eq "SharePoint"))
        {
         $PrintFlag = "False"
        }

        if(($SiteUrl.Length -ne 0) -and ($SiteUrl -ne $Site))
        {
         $PrintFlag = "False"
        }

        if (($FilterSites.Count -gt 0) -and (-not ($FilterSites -contains $Site)))
        {
         $PrintFlag = "False"
        }

        if($PrintFlag -eq "True")
        {
            #$ResultCount++
            $ActivityTime = (Get-Date($AuditData.CreationTime)).ToLocalTime()  
            $AccessFolder = $AuditData.SourceFileName
            $FolderURL = $AuditData.ObjectID
            
            $OutputEvents++
            $ExportResult = @{'Activity Time' = $ActivityTime; 'Folder Name' = $AccessFolder; 'Activity' = $Operation; 'Performed By' = $ActionBy; 'Folder URL' = $FolderURL;'Site URL' = $Site;'Workload' = $Workload;'More Info' = $MoreInfo}
            $ExportResults = New-Object PSObject -Property $ExportResult  
            $ExportResults | Sort 'Activity Time'| Select-Object 'Activity Time','Activity','Folder Name','Performed By','Folder Url','Site url','Workload','More Info' | Export-Csv -Path $OutputCSV -NoTypeInformation -Append 
        }
    }

    $currentResultCount=$CurrentResultCount+$ResultCount
    Write-Progress -Activity "`n     Retrieving folder activity audit log for $CurrentStart : $CurrentResultCount records"`n" Processed audit record count: $AggregateResultCount"

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

Write-Host "`n~~ Script prepared by AdminDroid Community ~~" -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green

# Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
