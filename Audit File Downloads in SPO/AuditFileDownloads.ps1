<#
=============================================================================================
Name:           Audit File Downloads in SharePoint Online using PowerShell
Version:        1.0
Website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script can generate 10+ file download audit reports.     
2. The script can be executed with an MFA enabled account too.     
3. The script retrieves file downloaded audit log for 180 days, by default.
4. Allows you audit file download report for a custom period. 
5. Helps to find recently downloaded files. e.g., Files downloaded in the last 30 days. 
6. Exports report results to CSV file.     
7. Identifies files downloaded by external/guest users. 
8. Helps to monitor all files downloaded by a specific user.
9. Helps to track SharePoint Online file downloads alone
10.Tracks OneDrive file downloads alone.
11.Finds files downloaded by a specific person in the specific period   
12. Automatically installs the EXO module (if not installed already) upon your confirmation.   
13. The script is scheduler friendly. i.e., Credentials can be passed as a parameter instead of saved inside the script. 
14. The script supports Certificate-based Authentication (CBA) too.

For detailed script execution: https://o365reports.com/2024/02/27/audit-file-downloads-in-sharepoint-online-using-powershell/
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [Int]$RecentlyDownloadedFiles_In_Days,
    [Switch]$SharePointOnlineOnly,
    [Switch]$OneDriveOnly,
    [switch]$FileDownloadedByExternalUsersOnly,
    [String]$DownloadedBy,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password
)

Function Connect_Exo
{
 #Check for EXO module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
   Import-Module ExchangeOnlineManagement
  } 
  else 
  { 
   Write-Host EXO module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 
 Write-Host Connecting to Exchange Online...
 #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential
 }
 elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
   Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
 }
 else
 {
  Connect-ExchangeOnline
 }
}

$MaxStartDate=((Get-Date).AddDays(-179)).Date

if($RecentlyDownloadedFiles_In_Days -ne "")
{
 $StartDate=((Get-Date).AddDays(-$RecentlyDownloadedFiles_In_Days)).Date
 $EndDate=(Get-Date).Date
}
#Retrieving Audit log for the past 180 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
#Getting start date for audit report
While($true)
{
 if ($StartDate -eq $null)
 {
  $StartDate=Read-Host Enter start time for report generation '(Eg:12/15/2023)'
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
   Write-Host `nAudit can be retrieved only for the past 180 days. Please select a date after $MaxStartDate -ForegroundColor Red
   return
  }
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}


#Getting end date to retrieve audit log
While($true)
{
 if ($EndDate -eq $null)
 {
  $EndDate=Read-Host Enter End time for report generation '(Eg: 12/15/2023)'
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
$Location=Get-Location
$OutputCSV="$Location\File_Download_Audit_Report$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
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

Connect_EXO
$AggregateResults = @()
$CurrentResult= @()
$CurrentResultCount=0
$AggregateResultCount=0
Write-Host `nRetrieving file download audit log from $StartDate to $EndDate... -ForegroundColor Yellow
$i=0
$ExportResult=""   
$ExportResults=@()  
$Operations="FileDownloaded"

if($FileDownloadedByExternalUsersOnly.IsPresent)
{
 $UserIds="*#EXT*"
}
elseif($DownloadedBy -ne "")
{
 $UserIds=$DownloadedBy
}
else
{
 $UserIds="*"
}

while($true)
{ 
 #Getting audit data for the given time range
 $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $Operations -UserIds $UserIds -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 $ResultCount=($Results | Measure-Object).count
 $AllAuditData=@()
 foreach($Result in $Results)
 {
  $i++
  $PrintFlag="True"
  $MoreInfo=$Result.auditdata
  $AuditData=$Result.auditdata | ConvertFrom-Json
  $ActivityTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()  #Get-Date($AuditData.CreationTime) Uncomment to view the Activity Time in UTC
  $UserID=$AuditData.userId
  $AccessedFile=$AuditData.SourceFileName
  $FileExtension=$AuditData.SourceFileExtension
  $SiteURL=$AuditData.SiteURL
  $Workload=$AuditData.Workload

  if($SharePointOnlineOnly.IsPresent -and ($Workload -ne "SharePoint"))
  {
   $PrintFlag="False"
  }

  if($OneDriveOnly.IsPresent -and ($Workload -ne "OneDrive"))
  {
   $PrintFlag="False"
  }

  #Export result to csv
  if($PrintFlag -eq "True")
  {
   $OutputEvents++
  $ExportResult=@{'Downloaded Time'=$ActivityTime;'Downloaded By'=$UserID;'Workload'=$Workload;'More Info'=$MoreInfo;'Downloaded File'=$AccessedFile;'Site URL'=$SiteURL;'File Extension'=$FileExtension}
  $ExportResults= New-Object PSObject -Property $ExportResult  
  $ExportResults | Select-Object 'Downloaded Time','Downloaded By','Downloaded File','Site URL','File Extension','Workload','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
 }
 }
 Write-Progress -Activity "`n     Retrieving file download audit data from $StartDate to $EndDate.."`n" Processed audit record count: $i"
 $currentResultCount=$CurrentResultCount+$ResultCount
 
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
 Write-Host `nThe exported report contains $OutputEvents audit records
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host `n "The Output file availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCSV" `n 
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$OutputCSV"   
  } 
 }
}
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue