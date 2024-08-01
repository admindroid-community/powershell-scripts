<#
=============================================================================================
Name:Track File activities in SharePoint Online Using PowerShell
Version: 1.0 
Website: o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the Exchange Online PowerShell module (if not installed already) upon your confirmation. 
2. Exports the file & folder usage report for the past 180 days into a CSV file. 
3. Allows to track file usage for a specific date range. 
4. Retrieves file activity by a specific user in the organization. 
5. Retrieves file activities by external or guest users. 
6. Allows to get file activities in SharePoint Online and OneDrive separately. 
7. Allows you to get weekly or monthly usage reports effortlessly. 
8. The script can be executed with an MFA-enabled account too. 
9. The script supports Certificate-based authentication (CBA). 
10. The script is scheduler friendly.

For detailed script execution: https://o365reports.com/2024/07/31/track-file-activities-in-sharepoint-online-using-powershell/
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$PerformedBy,
    [switch]$SharePointOnline,
    [switch]$OneDrive,
    [string]$UserName,
    [string]$Password,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbPrint    
)

Function Connect_Module {
    #Check for Exchange Online module installation
    $ExchangeModule = Get-Module ExchangeOnlineManagement -ListAvailable
    if($ExchangeModule.count -eq 0) {
        Write-Host ExchangeOnline module is not available -ForegroundColor Yellow
        $confirm = Read-Host Do you want to Install ExchangeOnline module? [Y] Yes  [N] No
        if($confirm -match "[Yy]") {
            Write-Host "Installing ExchangeOnline module ..."
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
            Import-Module ExchangeOnlineManagement
        }    
        else {
            Write-Host "ExchangeOnline Module is required. To Install ExchangeOnline module use 'Install-Module ExchangeOnlineManagement' cmdlet."
            Exit
        }
    }

    #Connect Exchange Online module
    Write-Host "`nConnecting Exchange Online module ..." -ForegroundColor Yellow
    if(($UserName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
        Connect-ExchangeOnline -Credential $Credential 
    }
    elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "") {
        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false
    
    }
}

$MaxStartDate=((Get-Date).AddDays(-180)).Date

#Getting file usage related activity for past 180 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
#Getting start date to generate file usage report
While($true)
{
 if ($StartDate -eq $null)
 {
  $StartDate=Read-Host Enter start time for report generation '(Eg:04/28/2021)'
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


#Getting end date to audit file usage report
While($true)
{
 if ($EndDate -eq $null)
 {
  $EndDate=Read-Host Enter End time for report generation '(Eg: 04/28/2021)'
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

$OutputCSV=".\FileUsageAuditReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
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

Connect_Module
$CurrentResultCount=0
$AggregateResultCount=0
Write-Host `nAuditing file usage from $StartDate to $EndDate...
$ProcessedAuditCount=0
$OutputEvents=0
$ExportResult=""   
$ExportResults=@()
$recordtype = "SharePointFileOperation"  

while($true)
{ 
 #Getting file usage audit data for the given time range
 Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -RecordType $recordtype -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 | ForEach-Object {
  $ResultCount++
  $ProcessedAuditCount++
  Write-Progress -Activity "`n     Retrieving file usage activities from $CurrentStart to $CurrentEnd.."`n" Processed audit record count: $ProcessedAuditCount"
  $MoreInfo=$_.auditdata
  $Operation=$_.Operations
  $ActionBy=$_.UserIds
  $AuditData=$_.auditdata | ConvertFrom-Json
  $Workload=$AuditData.Workload
  $PrintFlag="True"

  #Audit files usage by specific user
  if(($PerformedBy.Length -ne 0) -and ($PerformedBy -ne $ActionBy))
  {
   $PrintFlag="False"
  }

  #Filter for workload based file usage events
  If($SharePointOnline.IsPresent -and ($Workload -eq "OneDrive"))
  {
   $PrintFlag="False"
  }

  If($OneDrive.IsPresent -and ($Workload -eq "SharePoint"))
  {
   $PrintFlag="False"
  }

  if($PrintFlag -eq "True")
  {
   $ActivityTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()  #Get-Date($AuditData.CreationTime) Uncomment to view the Activity Time in UTC
   
   $AccessFile=$AuditData.SourceFileName
   $FileExtension=$AuditData.SourceFileExtension
   $Fileurl=$AuditData.ObjectID
   $SiteUrl=$AuditData.SiteUrl
   #Export result to csv
   $OutputEvents++
   $ExportResult=@{'Activity Time'=$ActivityTime;'File Name'=$AccessFile;'Activity'=$Operation;'Performed By'=$ActionBy;'File Extension'=$FileExtension;'File URL'=$Fileurl;'Site URL'=$Siteurl;'Workload'=$Workload;'More Info'=$MoreInfo}
   $ExportResults= New-Object PSObject -Property $ExportResult  
   $ExportResults | Select-Object 'Activity Time','Activity','File Name','Performed By','File Extension','File Url','Site url','Workload','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
  }
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
 Write-Host `nThe output file contains $OutputEvents audit records
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host `n" The Output file available in"  -NoNewline -ForegroundColor Yellow; Write-Host $OutputCSV 
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green  
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
  Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$OutputCSV"   
  } 
 }
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue