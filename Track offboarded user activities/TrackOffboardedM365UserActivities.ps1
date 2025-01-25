<#
=============================================================================================
Name:           Track Offboarded User Activities in Microsoft 365 using PowerShell 
Version:        1.0
website:        o365reports.com

Script Highlights:
~~~~~~~~~~~~~~~~~
1.The script uses modern authentication to connect to Exchange Online.  
2.The script can be executed with MFA enabled account too.  
3.Exports report results to CSV file.  
4.The script exports audit log for 180 days by default.
5.Allows you to keep audit log report for a custom period.  
6.Automatically installs the EXO module (if not installed already) upon your confirmation. 
7.The script is scheduler friendly. I.e., Credential can be passed as a parameter instead of saving inside the script. 
8.The script supports Certificate-based authentication (CBA).

For detailed Script execution: https://o365reports.com/2024/01/03/track-microsoft-365-offboarded-user-activities-using-powershell/
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserID,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$AdminName,
    [string]$Password
)

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


 $MaxStartDate=((Get-Date).AddDays(-179)).Date


#Retrive audit log for the past 180 days
if(($null -eq $StartDate) -and ($null -eq $EndDate))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
#Getting start date to audit export report
While($true)
{
 if ($null -eq $StartDate)
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


#Getting end date to export audit report
While($true)
{
 if ($null -eq $EndDate)
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


$OutputCSV=".\$UserId"+"_ActivityLogReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Check whether CurrentEnd exceeds EndDate
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

if($UserID -eq "")
{ write-host !!!!!
 $UserID=Read-Host Enter user UPN '(eg:John@contoso.com)'
}
Write-host ~~~~~~~~
$CurrentResultCount=0
$AggregateResultCount=0
Write-Host `nRetrieving user activity log from $StartDate to $EndDate... -ForegroundColor Yellow
$Count=0
$ExportResult=""   
$ExportResults=@()  
while($true)
{ 
 if($CurrentStart -eq $CurrentEnd)
 {
  Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
  Exit
 }
 #Getting audit log for given time range
 $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -UserIds $UserID -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 $ResultCount=($Results | Measure-Object).count
 foreach($Result in $Results)
 {
  $Count++
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
 Write-Progress -Activity "`n     Retrieving audit log from $StartDate to $EndDate.."`n" Processed audit record count: $Count"
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
