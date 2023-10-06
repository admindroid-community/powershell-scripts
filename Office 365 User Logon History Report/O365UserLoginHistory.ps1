<#
=============================================================================================
Name:           Office 365 User Login History Report
Website:        o365reports.com
Version:        3.0
For detailed Script execution: https://o365reports.com/2019/12/23/export-office-365-users-logon-history-report/
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$Success,
    [switch]$Failed,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserName,
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
 $StartDate=(((Get-Date).AddDays(-90))).Date
 $EndDate=Get-Date
}
else
{
 $StartDate=[DateTime]$StartDate
 $EndDate=[DateTime]$EndDate
 if($StartDate -lt ((Get-Date).AddDays(-90)))
 { 
  Write-Host `nAudit log can be retrieved only for past 90 days. Please select a date after (Get-Date).AddDays(-90) -ForegroundColor Red
  Exit
 }
 if($EndDate -lt ($StartDate))
 {
  Write-Host `nEnd time should be later than start time -ForegroundColor Red
  Exit
 }
}


 #Check for EXO v2 module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell V2 module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
   Import-Module ExchangeOnlineManagement
  } 
  else 
  { 
   Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 
 Write-Host Connecting to Exchange Online...
 #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
 if(($AdminName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $AdminName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential
 }
 else
 {
  Connect-ExchangeOnline
 }

$OutputCSV=".\AuditLog_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Filter for successful login attempts
if($success.IsPresent)
{
 $Operation="UserLoggedIn,TeamsSessionStarted,MailboxLogin"
}
#Filter for successful login attempts
elseif($Failed.IsPresent)
{
 $Operation="UserLoginFailed"
}
else
{
 $Operation="UserLoggedIn,UserLoginFailed,TeamsSessionStarted,MailboxLogin"
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
 #$Results.count
 $AllAuditData=@()
 $AllAudits=
 foreach($Result in $Results)
 {
  $AuditData=$Result.auditdata | ConvertFrom-Json
  $AuditData.CreationTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()
  $AllAudits=@{'Login Time'=$AuditData.CreationTime;'User Name'=$AuditData.UserId;'IP Address'=$AuditData.ClientIP;'Operation'=$AuditData.Operation;'Result Status'=$AuditData.ResultStatus;'Workload'=$AuditData.Workload}
  $AllAuditData= New-Object PSObject -Property $AllAudits
  $AllAuditData | Sort 'Login Time','User Name' | select 'Login Time','User Name','IP Address',Operation,'Result Status',Workload | Export-Csv $OutputCSV -NoTypeInformation -Append
 }
 Write-Progress -Activity "`n     Retrieving audit log from $StartDate to $EndDate.."`n" Processed audit record count: $AggregateResults"
 #$CurrentResult += $Results
 $currentResultCount=$CurrentResultCount+($Results.count)
 $AggregateResults +=$Results.count
 if(($CurrentResultCount -eq 50000) -or ($Results.count -lt 5000))
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
}

If($AggregateResults -eq 0)
{
 Write-Host No records found
}
else
{
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host `nThe Output file availble in $OutputCSV -ForegroundColor Green
 }
 Write-Host `nThe output file contains $AggregateResults audit records
}
