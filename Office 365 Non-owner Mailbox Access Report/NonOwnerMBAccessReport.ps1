Param
(
    [Parameter(Mandatory = $false)]
    [switch]$MFA,
    [Boolean]$IncludeExternalAccess=$false,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserName,
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


#Authentication using MFA
if($MFA.IsPresent)
{
 $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
 If ($MFAExchangeModule -eq $null)
 {
  Write-Host  `nPlease install Exchange Online MFA Module.  -ForegroundColor yellow
  Write-Host You can manually install module using below blog : `nhttps://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/ `nOR you can install module directly by entering "Y"`n
  $Confirm= Read-Host `nAre you sure you want to install module directly? [Y] Yes [N] No
  if($Confirm -match "[Y]")
  {
   Start-Process "iexplore.exe" "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"
  }
  else
  {
   Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
   Exit
  }
  $Confirmation= Read-Host Have you installed Exchange Online MFA Module? [Y] Yes [N] No
  if($Confirmation -match "[yY]")
  {
   $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
   If ($MFAExchangeModule -eq $null)
   {
    Write-Host Exchange Online MFA module is not available -ForegroundColor red
    Exit
   }
  }
  else
  {
   Write-Host Exchange Online PowerShell Module is required
   Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
   Exit
  }
 }

 #Importing Exchange MFA Module
 . "$MFAExchangeModule"
 Write-Host Enter credential in prompt to connect to Exchange Online
 Connect-EXOPSSession -WarningAction SilentlyContinue
}

#Authentication using non-MFA
else
{
 #Storing credential in script for scheduling purpose/ Passing credential as parameter
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
 }
 else
 {
  $Credential=Get-Credential -Credential $null
 }
 $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
 Import-PSSession $Session -AllowClobber -DisableNameChecking | Out-Null
}

$OutputCSV=".\NonOwner-Mailbox-Access-Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
$Operation='ApplyRecord','Copy','Create','FolderBind','HardDelete','MessageBind','Move','MoveToDeletedItem','RecordDelete','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules'


#Check whether CurrentEnd exceeds EndDate(checks for 1st iteration)
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

$AggregateResults = 0
$CurrentResult= @()
$CurrentResultCount=0
$NonOwnerAccess=0
Write-Host `nRetrieving audit log from $StartDate to $EndDate... -ForegroundColor Yellow

while($true)
{
 #Write-Host Retrieving audit log between StartDate $CurrentStart to EndDate $CurrentEnd ******* IntervalTime $IntervalTimeInMinutes minutes
 if($CurrentStart -eq $CurrentEnd)
 {
  Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
  Exit
 }


 #Getting Non-Owner mailbox access for a given time range
 else
 {
  $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $Operation -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 }
 $AllAuditData=@()
 $AllAudits=
 foreach($Result in $Results)
 {
  $AuditData=$Result.auditdata | ConvertFrom-Json

  #Remove owner access
  if($AuditData.LogonType -eq 0)
  {
   continue
  }

  #Filter for external access
  if(($IncludeExternalAccess -eq $false) -and ($AuditData.ExternalAccess -eq $true))
  {
   continue
  }

  #Processing non-owner mailbox access records
  if(($AuditData.LogonUserSId -ne $AuditData.MailboxOwnerSid) -or ((($AuditData.Operation -eq "SendAs") -or ($AuditData.Operation -eq "SendOnBehalf")) -and ($AuditData.UserType -eq 0)))
  {
   $AuditData.CreationTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()
   if($AuditData.LogonType -eq 1)
   {
    $LogonType="Administrator"
   }
   elseif($AuditData.LogonType -eq 2)
   {
    $LogonType="Delegated"
   }
   else
   {
    $LogonType="Microsoft datacenter"
   }
   if($AuditData.Operation -eq "SendAs")
   {
    $AccessedMB=$AuditData.SendAsUserSMTP
    $AccessedBy=$AuditData.MailboxOwnerUPN
   }
   elseif($AuditData.Operation -eq "SendOnBehalf")
   {
    $AccessedMB=$AuditData.SendOnBehalfOfUserSmtp
    $AccessedBy=$AuditData.MailboxOwnerUPN
   }
   else
   {
    $AccessedMB=$AuditData.MailboxOwnerUPN
    $AccessedBy=$AuditData.UserId
   }
   if($AccessedMB -eq $AccessedBy)
   {
    Continue
   }
  $NonOwnerAccess++
  $AllAudits=@{'Access Time'=$AuditData.CreationTime;'Accessed by'=$AccessedBy;'Performed Operation'=$AuditData.Operation;'Accessed Mailbox'=$AccessedMB;'Logon Type'=$LogonType;'Result Status'=$AuditData.ResultStatus;'External Access'=$AuditData.ExternalAccess}
  $AllAuditData= New-Object PSObject -Property $AllAudits
  $AllAuditData | Sort 'Access Time','Accessed by' | select 'Access Time','Logon Type','Accessed by','Performed Operation','Accessed Mailbox','Result Status','External Access' | Export-Csv $OutputCSV -NoTypeInformation -Append
 }
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
#Open output file after execution
else
{
 Write-Host `nThe output file contains $NonOwnerAccess audit records
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
