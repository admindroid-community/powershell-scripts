﻿<#
=============================================================================================
Name:           Audit SharePoint Online group membership changes report
Website:        o365reports.com
For detailed script execution:  https://o365reports.com/2022/06/22/audit-sharepoint-online-group-membership-changes-using-powershell
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [switch]$SharePointOnly,
    [switch]$OneDriveOnly,
    [switch]$GuestOnly,
    [string]$UserName,
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
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential
 }
 else
 {
  Connect-ExchangeOnline
 }
}

$MaxStartDate=((Get-Date).AddDays(-89)).Date

#Audit SharePoint group membership changes for past 90 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
#Getting start date to audit SharePoint group membership changes report
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
   Write-Host `nAudit can be retrieved only for past 90 days. Please select a date after $MaxStartDate -ForegroundColor Red
   return
  }
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}


#Getting end date to audit SPO group membership changes report
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

$OutputCSV=".\AuditSPOGrpMembershipChanges_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
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
$CurrentResultCount=0
$AggregateResultCount=0
Write-Host `nAuditing SharePoint group membership changes from $StartDate to $EndDate...
$ProcessedAuditCount=0
$OutputEvents=0
$ExportResult=""   
$ExportResults=@()  


while($true)
{
 #Getting audit data for the given time range
 Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations "AddedtoGroup,RemovedfromGroup" -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 | ForEach-Object {
  $ResultCount++
  $ProcessedAuditCount++
  Write-Progress -Activity "`n     Retrieving SharePoint Group membership changes audit from $CurrentStart to $CurrentEnd.."`n" Processed audit record count: $ProcessedAuditCount"
  $MoreInfo=$_.auditdata
  $Operation=$_.Operations
  $User=$_.UserIds
  $PrintFlag="True"
  $AuditData=$_.auditdata | ConvertFrom-Json
  $SiteUrl=$AuditData.SiteUrl
  $MemberName=$AuditData.TargetUserorGroupName
  $MemberType=$AuditData.TargetUserorGroupType
  $Html=$AuditData.EventData
  $Group = [Regex]::new("(?<=Group>)(.*)(?=</Group>)")
  $AllowShare=[Regex]::new("(?<=ShareApplied>)(.*)(?=</MembersCan)")
  $MembersCanShare=$AllowShare.Match($Html).Value
  if($MembersCanShare -eq "")
  {$MembersCanShare="-"}
  $GroupName = $Group.Match($html).Value  
  $Result=$AuditData.ResultStatus
  $Workload=$AuditData.Workload

  #Filter result based on Workload
  if(($SharePointOnly.IsPresent) -and ($Workload -eq "OneDrive"))
  {
   $PrintFlag="False"
  }
  elseif(($OneDriveOnly.IsPresent) -and ($Workload -eq "SharePoint"))
  {
   $PrintFlag="False"
  }

  #Filter Guest users alone
  if(($GuestOnly.IsPresent) -and ($MemberType -ne "Guest"))
  {
   $PrintFlag="False"
  }  
 
  $EventTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()  #Get-Date($AuditData.CreationTime) Uncomment to view the Activity Time in UTC
  #Export result to csv
  if($PrintFlag -eq "True")
  {
   $OutputEvents++
   $ExportResult=@{'Event Time'=$EventTime;'Performed By'=$User; 'Operation'=$Operation;'Site URL'=$SiteUrl;'Member Name'=$MemberName;'Member Type'=$MemberType;'Workload'=$Workload;'Group Name'=$GroupName;'Members Can Share Sites&Files'=$MembersCanShare;'More Info'=$MoreInfo}
   $ExportResults= New-Object PSObject -Property $ExportResult  
   $ExportResults | Select-Object 'Event Time','Performed By','Operation','Group Name','Member Name','Member Type','Members Can Share Sites&Files','Site URL','Workload','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
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
  Write-Host `nThe Output file available in $OutputCSV -ForegroundColor Green
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