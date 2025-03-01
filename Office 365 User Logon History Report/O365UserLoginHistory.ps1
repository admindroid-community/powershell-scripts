<#
=============================================================================================
Name:           Office 365 User Login History Report
Website:        o365reports.com
Version:        4.0

Script Highlights: 
~~~~~~~~~~~~~~~~~

1.The script automatically installs the EXO PowerShell module (if not installed already) upon your confirmation.
2.Allows you to filter the result based on successful and failed logon attempts. 
3.The exported report has IP addresses from where your office 365 users are login. 
4.This script can be executed with MFA enabled account. 
5.You can export the report to choose either “All Office 365 users’ login attempts” or “Specific Office user’s logon attempts”. 
6.By using advanced filtering options, you can export “Office 365 users Sign-in report” and “Suspicious login report”. 
7.Exports report result to CSV. 
8.Helps to track workload based sign-in history, such as Entra, Exchange Online, SharePoint Online, MS Teams
9.This script is scheduler friendly. I.e., credentials can be passed as a parameter instead of saving inside the script. 
10.Supports certificate-based authentication too.

For detailed Script execution: https://o365reports.com/2019/12/23/export-office-365-users-logon-history-report/




Change Log
~~~~~~~~~~

    V1.0 (Dec 23, 2019) - File created
    V2.0 (Aug 10, 2022) - Upgraded from Exchange Online PowerShell V1 module.
    V2.1 (Oct 06, 2023) - Minor usability improvements.
    V3.0 (May 24, 2024) - Added certificate-based authentication support to enhance scheduling capability.
    V4.0 (Mar 01, 2025) - Added workload param to enhance filtering capability.
============================================================================================
#>


Param
(
    [Parameter(Mandatory = $false)]
    [switch]$Success,
    [switch]$Failed,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [ValidateSet(
        "EntraID", 
        "MicrosoftTeams",
        "Exchange", 
        "SharePoint"
    )]
    [string[]]$Workload,
    [string]$UserName,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$AdminName,
    [string]$Password
)

#Getting StartDate and EndDate for Audit log
if ((($StartDate -eq $null) -and ($EndDate -ne $null)) -or (($StartDate -ne $null) -and ($EndDate -eq $null)))
{
 Write-Host `nPlease enter both StartDate and EndDate for Audit log collection -ForegroundColor Red
 exit
}   
elseif(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $StartDate=(((Get-Date).AddDays(-180))).Date
 $EndDate=Get-Date
}
else
{
 $StartDate=[DateTime]$StartDate
 $EndDate=[DateTime]$EndDate
 if($StartDate -lt ((Get-Date).AddDays(-180)))
 { 
  Write-Host `nAudit log can be retrieved only for past 180 days. Please select a date after (Get-Date).AddDays(-180) -ForegroundColor Red
  Exit
 }
 if($EndDate -lt ($StartDate))
 {
  Write-Host `nEnd time should be later than start time -ForegroundColor Red
  Exit
 }
}

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
 if(($AdminName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $AdminName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
 }
 elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
   Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
 }
 else
 {
  Connect-ExchangeOnline -ShowBanner:$false
 }

# Map friendly names to actual operations
$WorkloadOperations = @{
    "EntraID" = "UserLoggedIn,UserLoginFailed";
    "MicrosoftTeams" = "TeamsSessionStarted";
    "Exchange" = "MailboxLogin";
    "SharePoint" = "SignInEvent"
}


$Location=Get-Location
$OutputCSV="$Location\UserLoginHistoryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Apply filters based on params
if ($Failed.IsPresent) {
 $Operation="UserLoginFailed"
} elseif (-not [string]::IsNullOrEmpty($Workload)) {
 $Operation = ($Workload | ForEach-Object { $WorkloadOperations[$_] }) -join ","
} elseif ($Success.IsPresent) {
 $Operation="UserLoggedIn,TeamsSessionStarted,MailboxLogin,SignInEvent"
} else {
 $Operation="UserLoggedIn,UserLoginFailed,TeamsSessionStarted,MailboxLogin,SignInEvent"
}

#Check whether CurrentEnd exceeds EndDate(checks for 1st iteration)
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

$AggregateResults = 0
$CurrentResult= @()
$CurrentResultCount=0
Write-Host `nRetrieving audit log from $StartDate to $EndDate... -ForegroundColor Yellow

while($true)
{ 
 #Write-Host Retrieving audit log between StartDate $CurrentStart to EndDate $CurrentEnd ******* IntervalTime $IntervalTimeInMinutes minutes
 if($CurrentStart -eq $CurrentEnd)
 {
  Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
  Exit
 }

 #Getting audit log for specific user(s) for a given time range
 if($UserName -ne "")
 {
  $Results=Search-UnifiedAuditLog -UserIds $UserName -StartDate $CurrentStart -EndDate $CurrentEnd -operations $Operation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 }

 #Getting audit log for all users for a given time range
 else
 {
  $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $Operation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 }
 $ResultsCount=($Results|Measure-Object).count
 $AllAuditData=@()
 $AllAudits=
 foreach($Result in $Results)
 {
  $AuditData=$Result.auditdata | ConvertFrom-Json
  $AuditData.CreationTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()
  $AllAudits=@{'Login Time'=$AuditData.CreationTime;'User Name'=$AuditData.UserId;'IP Address'=$AuditData.ClientIP;'Operation'=$AuditData.Operation;'Result Status'=$AuditData.ResultStatus;'Workload'=$AuditData.Workload;}
  $AllAuditData= New-Object PSObject -Property $AllAudits
  $AllAuditData | Sort 'Login Time','User Name' | select 'Login Time','User Name','IP Address',Operation,'Result Status',Workload | Export-Csv $OutputCSV -NoTypeInformation -Append
 }
 
 #$CurrentResult += $Results
 $currentResultCount=$CurrentResultCount+$ResultsCount
 $AggregateResults +=$ResultsCount
 Write-Progress -Activity "`n     Retrieving audit log from $CurrentStart to $CurrentEnd.."`n" Processed audit record count: $AggregateResults"
 if(($CurrentResultCount -eq 50000) -or ($ResultsCount -lt 5000))
 {
  if($CurrentResultCount -eq 50000)
  {
   Write-Host Retrieved max record for the current range.Proceeding further may cause data loss or rerun the script with reduced time interval. -ForegroundColor Red
   $Confirm=Read-Host `nAre you sure you want to continue? [Y] Yes [N] No
   if($Confirm -notmatch "[Y]")
   {
    Write-Host Please rerun the script with reduced time interval -ForegroundColor Red
    Exit
   }
   else
   {
    Write-Host Proceeding audit log collection with data loss
   }
  } 
  #Check for last iteration
  if(($CurrentEnd -eq $EndDate))
  {
   break
  }
  [DateTime]$CurrentStart=$CurrentEnd
  #Break loop if start date exceeds current date(There will be no data)
  if($CurrentStart -gt (Get-Date))
  {
   break
  }
  [DateTime]$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
  if($CurrentEnd -gt $EndDate)
  {
   $CurrentEnd=$EndDate
  }
  
  $CurrentResultCount=0
  $CurrentResult = @()
 }
 $c=($Results | Measure-Object).Count
}

#Open output file after execution
If($AggregateResults -eq 0)
{
 Write-Host No records found
}
else
{
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host ""
  Write-Host " The Output file availble in:" -NoNewline -ForegroundColor Yellow
  Write-Host $OutputCSV 
    Write-Host `nThe output file contains $AggregateResults audit records
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n 
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
