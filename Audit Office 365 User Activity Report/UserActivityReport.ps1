<#
=============================================================================================
Name:           Export Office 365 User Activity Report to CSV using PowerShell 
Version:        2.0
website:        o365reports.com

Script Highlights:
~~~~~~~~~~~~~~~~~
1.The script uses modern authentication to connect to Exchange Online.  
2.The script can be executed with MFA enabled account too.  
3.Exports report results to CSV file.  
4.Allows you to generate a user activity report for a custom period.  
5.Automatically installs the EXO V2 module (if not installed already) upon your confirmation. 
6.The script is scheduler friendly. I.e., Credential can be passed as a parameter instead of saving inside the script. 

For detailed Script execution: https://o365reports.com/2021/01/06/export-office-365-user-activity-report-to-csv-using-powershell/
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [switch]$MFA,
    [switch]$Default,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserID,
    [string]$AdminName,
    [string]$Password
)

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
 } 
 else 
 { 
  Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
  Exit
 }
} 

#Check for MSOnline module 
$Module=Get-Module -Name MSOnline -ListAvailable  
if($Module.count -eq 0) 
{ 
 Write-Host MSOnline module is not available  -ForegroundColor yellow  
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
 if($Confirm -match "[yY]") 
 { 
  Install-Module MSOnline 
  Import-Module MSOnline
 } 
 else 
 { 
  Write-Host MSOnline module is required to connect AzureAD.Please install module using Install-Module MSOnline cmdlet. 
  Exit
 }
} 

#Connect Exchange Online with MFA
 if($MFA.IsPresent)
 {
  Write-Host Connecting to Exchange Online...
  Connect-ExchangeOnline
  Write-Host Connecting to MSOnline module...
  Connect-MsolService
 }

#Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
else
{
 if(($AdminName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $AdminName,$SecuredPassword
 }
 else
 {
  $Credential=Get-Credential -Credential $null
 }
 Write-Host Connecting to Exchange Online...
 Connect-ExchangeOnline -Credential $Credential
 Write-Host Connecting to MSonline module...
 Connect-MsolService -Credential $Credential
}

#Getting user activity for past 90 days
if($Default.IsPresent)
{
 $EndDate=(Get-Date).Date
 $StartDate= ((Get-Date).AddDays(-89)).Date
}
 
#Getting start date for Audit log  
While($true)
{
 if ($StartDate -eq $null)
 {
  $StartDate=Read-Host Enter start time for audit collection '(Eg:11/20/2019)'
 }
 Try
 {
  $Date=[DateTime]$StartDate
  if($Date -gt ((Get-Date).AddDays(-90)))
  { 
   break
  }
  else
  {
   Write-Host `nAudit log can be retrieved only for past 90 days. Please select a date after (Get-Date).AddDays(-90) -ForegroundColor Red
   return
  }
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}


#Getting end date for Audit log
While($true)
{
 if ($EndDate -eq $null)
 {
  $EndDate=Read-Host Enter End time for audit collecton '(Eg: 11/20/2019)'
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

#Checking whether the user is available
if((Get-MsolUser -UserPrincipalName $userID) -eq $null)
{
 Write-Host User does not exist. Please check the user name.
 exit
}

$OutputCSV=".\UserActivityReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Check whether CurrentEnd exceeds EndDate
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

$AggregateResults = @()
$CurrentResult= @()
$CurrentResultCount=0
$AggregateResultCount=0
Write-Host `nRetrieving user activity log from $StartDate to $EndDate... -ForegroundColor Yellow
$i=0
$ExportResult=""   
$ExportResults=@()  
while($true)
{ 
 #Write-Host Retrieving user activity log between StartDate $CurrentStart to EndDate $CurrentEnd ******* IntervalTime $IntervalTimeInMinutes minutes
 if($CurrentStart -eq $CurrentEnd)
 {
  Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
  Exit
 }
 #Write-Host !!!!!!!!!!!!!!
 #Getting audit log for given time range
 $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -UserIds $UserID -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 $ResultCount=($Results | Measure-Object).count
 $AllAuditData=@()
 foreach($Result in $Results)
 {
  $i++
  $MoreInfo=$Result.auditdata
  $AuditData=$Result.auditdata | ConvertFrom-Json
  $ActivityTime=Get-Date($AuditData.CreationTime) -format g
  $UserID=$AuditData.userId
  $Operation=$AuditData.Operation
  $ResultStatus=$AuditData.ResultStatus
  $Workload=$AuditData.Workload

  #Export result to csv
  $ExportResult=@{'Activity Time'=$ActivityTime;'User Name'=$UserID;'Operation'=$Operation;'Result'=$ResultStatus;'Workload'=$Workload;'More Info'=$MoreInfo}
  $ExportResults= New-Object PSObject -Property $ExportResult  
  $ExportResults | Select-Object 'Activity Time','User Name','Operation','Result','Workload','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
 }
 Write-Progress -Activity "`n     Retrieving audit log from $StartDate to $EndDate.."`n" Processed audit record count: $i"
 $currentResultCount=$CurrentResultCount+$ResultCount
 if($CurrentResultCount -eq 50000)
 {
  Write-Host Retrieved max record for current range.Proceeding further may cause data loss or rerun the script with reduced time interval. -ForegroundColor Red
  $Confirm=Read-Host `nAre you sure you want to continue? [Y] Yes [N] No
  if($Confirm -match "[Y]")
  {
   Write-Host Agg $AggregateResultCount CurrentResu $CurrentResultCount
   $AggregateResultCount +=$CurrentResultCount
   Write-Host Proceeding audit log collection with data loss
   [DateTime]$CurrentStart=$CurrentEnd
   [DateTime]$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
   $CurrentResultCount=0
   $CurrentResult = @()
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

 
 if($Results.count -lt 5000)
 {
  #$AggregateResults +=$CurrentResult
  $AggregateResultCount +=$CurrentResultCount
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
  $CurrentResult = @()
  if($CurrentEnd -gt $EndDate)
  {
   $CurrentEnd=$EndDate
  }
 }
}

If($AggregateResultCount -eq 0)
{
 Write-Host No records found
}
else
{
 Write-Host `nThe output file contains $AggregateResultCount audit records `n
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host " The Output file available in:" -NoNewline -ForegroundColor Yellow
  Write-Host $OutputCSV 
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
 Disconnect-ExchangeOnline -Confirm:$false | Out-Null
