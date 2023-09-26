<#
=============================================================================================
Name:           SharePoint Online Anonymous Link Activity Report
Description:    This script exports SharePoint Online anonymous link  activities report to CSV
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.Allow to generate 8 different anonymous link reports. 
2.The script uses modern authentication to retrieve audit logs.   
3.The script can be executed with MFA enabled account too.   
4.Exports report results to CSV file.   
5.Automatically installs the EXO V2 module (if not installed already) upon your confirmation.  
6.The script is scheduler friendly. I.e., Credential can be passed as a parameter instead of saving inside the script. 

For detailed script execution: https://o365reports.com/2021/06/22/audit-anonymous-access-in-sharepoint-online-using-powershell
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [switch]$SharePointOnline,
    [switch]$OneDrive,
    [switch]$AnonymousSharing,
    [switch]$AnonymousAccess,
    [string]$AdminName,
    [string]$Password
)

Function Connect_Exo
{
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
 Write-Host `nConnecting to Exchange Online...
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
}

$MaxStartDate=((Get-Date).AddDays(-89)).Date

#Getting anonymous activity for past 90 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
$startDate
#Getting start date to generate anonymous activity report
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
   Write-Host `nAnonymous activity report can be retrieved only for past 90 days. Please select a date after $MaxStartDate -ForegroundColor Red
   return
  }
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}


#Getting end date to generate external sharing report
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

$OutputCSV=".\AnonymousLinksActivityReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
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
Write-Host `nRetrieving anonymous link events from $StartDate to $EndDate...
$ProcessedAuditCount=0
$OutputEvents=0
$ExportResult=""   
$ExportResults=@()  
if($AnonymousSharing.IsPresent)
{
 $RetriveOperation="AnonymousLinkCreated"
}
elseif($AnonymousAccess.IsPresent)
{
 $RetriveOperation="AnonymousLinkUsed"
}
else
{
 $RetriveOperation="AnonymousLinkRemoved,AnonymousLinkcreated,AnonymousLinkUpdated,AnonymousLinkUsed"
}

while($true)
{ 
 #Getting anonymous sharing/access audit data for given time range
 $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $RetriveOperation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 $ResultCount=($Results | Measure-Object).count
 foreach($Result in $Results)
 {
  $ProcessedAuditCount++
  $MoreInfo=$Result.auditdata
  $Operation=$Result.Operations
  $AuditData=$Result.auditdata | ConvertFrom-Json
  $Workload=$AuditData.Workload

  #Filter for SharePointOnline external Sharing events
  If($SharePointOnline.IsPresent -and ($Workload -eq "OneDrive"))
  {
   continue
  }

  If($OneDrive.IsPresent -and ($Workload -eq "SharePoint"))
  {
   continue
  }
  
  $ActivityTime=Get-Date($AuditData.CreationTime) -format g
  $PerformedBy=$AuditData.userId
  $ResourceType=$AuditData.ItemType
  $Resource=$AuditData.ObjectId
  $SiteURL=$AuditData.SiteURL
  $UserIP=$AuditData.ClientIP
  $EventData=$AuditData.EventData
  
  #Check whether the anonymous link has edit permission
  if($Operation -ne "AnonymousLinkUsed")
  {
   if($EventData -like "*View*")
   {
    $EditEnabled="False"
   }
   else
   {
    $EditEnabled="True"
   }
  }
  else
  {
   $EditEnabled="NA"
  }

  #Export result to csv
  $OutputEvents++
  $ExportResult=@{'Activity Time'=$ActivityTime;'Activity'=$Operation;'Performed By'=$PerformedBy;'User IP'=$UserIP;'Resource Type'=$ResourceType;'Shared/Accessed Resource'=$Resource;'Site url'=$Siteurl;'Workload'=$Workload;'Edit Enabled'=$EditEnabled;'More Info'=$MoreInfo}
  $ExportResults= New-Object PSObject -Property $ExportResult  
  $ExportResults | Select-Object 'Activity Time','Activity','Performed By','User IP','Resource Type','Shared/Accessed Resource','Edit Enabled','Site URL','Workload','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
 }
 Write-Progress -Activity "`n     Retrieving anonymous link activities from $CurrentStart to $CurrentEnd.."`n" Processed audit record count: $ProcessedAuditCount"
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
  #$AggregateResultCount +=$CurrentResultCount
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

If($OutputEvents -eq 0)
{
 Write-Host No records found
}
else
{
 Write-Host `nThe output file contains $OutputEvents audit records
 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host `n The Output file availble in: -NoNewline -ForegroundColor Yellow
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
#Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue