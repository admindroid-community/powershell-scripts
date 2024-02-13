<#
=============================================================================================
Name:           Audit user creations in Microsoft 365
Version:        1.0
Website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script uses modern authentication to retrieve audit logs.     
2. The script can be executed with an MFA enabled account too.       
3. Exports report results to CSV file.     
4. Identifies who created Azure AD guest users. 
5. Helps to find recently created users. e.g., users created in the last 30 days. 
6. Allows you to generate a user creation audit report for a custom period.    
7. Automatically installs the EXO module (if not installed already) upon your confirmation.   
8. The script is scheduler friendly. i.e., Credentials can be passed as a parameter instead of saved inside the script. 

For detailed script execution:  https://o365reports.com/2023/08/01/find-who-created-user-account-in-microsoft-365
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [Int]$RecentlyCreatedDays,
    [switch]$GuestUsersOnly,
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
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
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
 else
 {
  Connect-ExchangeOnline
 }
}

$MaxStartDate=((Get-Date).AddDays(-89)).Date
if($RecentlyCreatedDays -ne "")
{
 $StartDate=((Get-Date).AddDays(-$RecentlyCreatedDays)).Date
 $EndDate=(Get-Date).Date
 $StartDate
 $EndDate
}

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

$OutputCSV=".\Audit_M365_User_Creations_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
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
Write-Host `nAuditing Microsoft 365 user creations from $StartDate to $EndDate... -ForegroundColor Cyan
$ProcessedAuditCount=0
$OutputEvents=0
$ExportResult=""   
$ExportResults=@()  
$Operations="Add user"


while($true)
{ 
 #Getting audit data for the given time range
 Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $Operations -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000 | ForEach-Object {
  $ResultCount++
  $ProcessedAuditCount++
  $PrintFlag=$true
  Write-Progress -Activity "`n     Retrieving Microsoft 365 user creations from $CurrentStart to $CurrentEnd.."`n" Processed audit record count: $ProcessedAuditCount"
  $MoreInfo=$_.auditdata
  $Operation=$_.Operations
  $CreatedBy=$_.UserIds
  $AuditData=$_.auditdata | ConvertFrom-Json
  $ResultStatus=$AuditData.ResultStatus
  $EventTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()  #Get-Date($AuditData.CreationTime) Uncomment to view the Activity Time in UTC
  $UserName=$AuditData.ObjectId
  $Properties=$AuditData.ModifiedProperties
  $AccountStatus=(($Properties | Where-Object { $_.Name -eq 'AccountEnabled' }).NewValue -replace '[\[\]]', '').Trim()
  $DisplayName=(($Properties | Where-Object { $_.Name -eq 'DisplayName' }).NewValue -replace '["\[\]]', '').Trim()
  $UserType=($Properties | Where-Object { $_.Name -eq 'UserType' }).NewValue -replace '["\[\]]', ''
  $UserType=$UserType.Trim()
  if($GuestUsersOnly.IsPresent -and $UserType -ne "Guest")
  {
   $PrintFlag=$false
  }
  $IncludedProperties=($Properties | where {$_.Name -eq "Included Updated Properties"}).NewValue
  $propertyNames = $IncludedProperties -split ", "
  if($AccountStatus -eq $true)
  {
   $AccountStatus="Enabled"
  }
  elseif($AccountStatus -eq $false)
  {
   $AccountStatus="Disabled"
  }

 $SkuName = "Created without license"

if ($propertyNames -contains "AssignedLicense") {
    $assignedLicenseProperty = $Properties | Where-Object { $_.Name -eq 'AssignedLicense' }
    $SkuNames = $assignedLicenseProperty.NewValue | Select-String -Pattern 'SkuName=([^,]+)' -AllMatches | ForEach-Object { $_.Matches.Value -replace 'SkuName=', '' }

    if ($SkuNames.Count -gt 0) {
        $SkuName = $SkuNames -join ', '
    }
}
  #Export result to csv
  if($PrintFlag -eq "True")
  {
   $OutputEvents++
   $ExportResult=@{'Creation Time'=$EventTime;'Created By'=$CreatedBy; 'UPN'=$UserName;'Display Name'=$DisplayName;'User Type'=$UserType;'Account Status'=$AccountStatus;'License'=$SkuName;'Result Status'=$ResultStatus;'More Info'=$MoreInfo}
   $ExportResults= New-Object PSObject -Property $ExportResult  
   $ExportResults | Select-Object 'Creation Time','Display Name','UPN','Created By','Account Status','License','User Type','Result Status','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
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