<#
=============================================================================================
Name:           Audit mailbox permission changes in Office 365
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.The script uses modern authentication to retrieve audit logs.   
2.The script can be executed with an MFA enabled account too.     
3.Exports report results to CSV file.     
4.Excludes system generated permission changes by default. 
5.Helps to detect who modified mailbox permissions.  
6.Tracks who granted full access, send as, send on behalf permissions separately. 
7.Allows you to generate mailbox permission changes audit report for a custom period.     
8.Automatically installs the EXO V2 module (if not installed already) upon your confirmation.    
9.The script is scheduler-friendly. i.e., Credential can be passed as a parameter instead of saving inside the script.

For detailed script execution:  https://o365reports.com/2022/07/06/audit-mailbox-permission-changes-in-office-365-powershell
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [switch]$SendAsOnly,
    [switch]$SendOnBehalfOnly,
    [switch]$FullAccessOnly,
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

#Audit mailbox permission changes for past 90 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
#Getting start date to audit mailbox permission changes report
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


#Getting end date to audit mailbox permission changes changes report
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

$OutputCSV=".\AuditMailboxPermissionChanges_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
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
Write-Host `nAuditing mailbox permission changes from $StartDate to $EndDate...
$ProcessedAuditCount=0
$OutputEvents=0
$ExportResult=""   
$ExportResults=@()  

#Determine the operations based on the user's input
If($SendAsOnly.IsPresent)
{
 $Operations="Add-RecipientPermission,Remove-RecipientPermission"
}
elseif($SendOnBehalfOnly.IsPresent)
{
 $Operations="Set-mailbox"
}
elseif($FullAccessOnly.IsPresent)
{
 $Operations="Add-MailboxPermission,Remove-MailboxPermission"
}
else
{
 $Operations="Add-RecipientPermission,Remove-RecipientPermission,Set-mailbox,Add-MailboxPermission,Remove-MailboxPermission"
}


while($true)
{
 #Getting audit data for the given time range
 Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $Operations -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 | ForEach-Object {
  $ResultCount++
  $ProcessedAuditCount++
  Write-Progress -Activity "`n     Retrieving mailbox permission changes from $CurrentStart to $CurrentEnd.."`n" Processed audit record count: $ProcessedAuditCount"
  $MoreInfo=$_.auditdata
  $Operation=$_.Operations
  $User=$_.UserIds
  $PrintFlag="True"
  if($user -like "NT Authority*")
  {
   $PrintFlag="False"
  } 
  $AuditData=$_.auditdata | ConvertFrom-Json
  $EventTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()  #Get-Date($AuditData.CreationTime) Uncomment to view the Activity Time in UTC
 
  #Check for sendonbehalf operation
  if(($Operation -eq "Set-Mailbox") -and ($auditData.Parameters.name -contains "GrantSendOnBehalfTo"))
  { 
   $MailboxInfo=$AuditData.Parameters.Value[0]
   $DelegateInfo=$AuditData.Parameters.Value[1]
   $MailboxName=($MailboxInfo -split "/") | Select-Object -Last 1
   if($DelegateInfo -like "*Microsoft Exchange Hosted*")
   {
    $Permission="Adding send on behalf permission"
    $Delegate= ($DelegateInfo -split "/") | Select-Object -Last 1
   }
   else
   {
    if($DelegateInfo -eq "")
   {
    $DelegateInfo="-"
   }
    $Permission="Removing send on behalf permission"
    $Delegate="Current Delegates:"+ $DelegateInfo
   }
  }
  elseif($Operation -eq "Set-Mailbox")
  {
   $PrintFlag="False"
  }
  else
  {
   $MailboxInfo=$AuditData.Parameters.Value[0]
   $DelegateInfo=$AuditData.Parameters.Value[1]
   $MailboxName=($MailboxInfo -split "/") | Select-Object -Last 1
   $Delegate=$Delegate= ($DelegateInfo -split "/") | Select-Object -Last 1
   if($Operation -eq "Add-RecipientPermission")
   {
    $Permission="Adding send as permission"
   }
   elseif($Operation -eq "Remove-RecipientPermission")
   {
    $Permission="Removing send as permission"
   }
   elseif($Operation -eq "Add-MailboxPermission")
   {
    $Permission="Adding full access permission"
   }
   elseif($Operation -eq "Remove-MailboxPermission")
   {
    $Permission="Removing full access permission"
   }
  }

  #Export result to csv
  if($PrintFlag -eq "True")
  {
   $OutputEvents++
   $ExportResult=@{'Event Time'=$EventTime;'Performed By'=$User; 'Permission'=$Permission;'Mailbox Name'=$MailboxName;'Delegate Name'=$Delegate;'More Info'=$MoreInfo}
   $ExportResults= New-Object PSObject -Property $ExportResult  
   $ExportResults | Select-Object 'Event Time','Permission','Performed By','Mailbox Name','Delegate Name','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
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
  Write-Host `n "The Output file available in:" -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCSV"`n
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$OutputCSV"   
  } 
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 }
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue