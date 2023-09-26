<#
=============================================================================================
Name:           Exchange Online Archive MAilbox Size Report
Description:    This script exports Exchange Online Archive Mailbox sizes to CSV
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.Validates and installs Exchange Online PowerShell module upon user confirmation (if not exists already) 
2.Facility to supply the input and export the output in the CSV format 
3.The report is delivered with the significant attributes. As it is the user–friendly script, you can publish desired attributes too. 
4.Supports both MFA and Non-MFA accounts
5.The script is scheduler-friendly. You can schedule the report generation upon giving UserName and Password. 

For detailed script execution: https://o365reports.com/2021/03/30/export-office-365-archive-mailbox-size-report-using-powershell
============================================================================================
#>

param
(
    [string] $UserName = $null,
    [string] $Password = $null,
    [Switch] $UserMBOnly,
    [Switch] $SharedMBOnly,
    [Switch] $AutoExpandingArchiveEnabled,
    [String] $MBIdentityFile
)

function CSVImport {
 $IdentityList = Import-Csv -Header "IdentityValue" $MBIdentityFile
 foreach($MailboxDetails in $IdentityList) {
  $currIdentity = $MailboxDetails.IdentityValue
  if($null -eq $WhereObjectCheck) 
  {
   $UserData = Get-Mailbox -Identity $currIdentity -Archive -ErrorAction SilentlyContinue
  }
  else 
  {
   $UserData = Get-Mailbox -Identity $currIdentity -Archive -ErrorAction SilentlyContinue | Where-Object $WhereObjectCheck
  }
  if($null -eq $UserData) 
  {
   Write-Host $currIdentity mailbox is not archive enabled/Invalid.
  }
  else 
  {
   ExportOutput
  }
 }
}
function ExportOutput {
 $ArchiveMailboxSize = ((Get-MailboxStatistics -Identity $UserData.UserPrincipalName -Archive -WarningAction SilentlyContinue).TotalItemSize)
 if($null -ne $ArchiveMailboxSize) 
 {
  $ArchiveMailboxSize = $ArchiveMailboxSize.ToString().split("()")
  $ArchiveMailboxSizeRounded = $ArchiveMailBoxSize | Select-Object -Index 0
  $ArchiveMailboxSizeBytes = ($ArchiveMailBoxSize | Select-Object -Index 1).ToString().Split(" ") | Select-Object -Index 0
 }
 else
 {
  $ArchiveMailboxSizeRounded = $ArchiveMailboxSizeBytes = "0"
 }
 if($UserData.AutoExpandingArchiveEnabled -eq $True)
 {
  $AutoExpandArchive = "Enabled"
 }
 else
 {
  $AutoExpandArchive = "Disabled" 
 }
 $CurrExportValue = $UserData.DisplayName
 $ArchiveQuotaSize = ($UserData.ArchiveQuota).ToString().split("()") | Select-Object -Index 0
 $ArchiveWarningQuotaSize = ($UserData.ArchiveWarningQuota).ToString().split("()") | Select-Object -Index 0

 $global:ReportSize = $global:ReportSize + 1
 Write-Progress -Activity "Exporting $CurrExportValue`n Processed Mailbox count: $global:ReportSize" -Status "Preparing Report"

 #Export output to CSV
 $ExportResult = @{'DisplayName' = $UserData.DisplayName; 'Email Address' = $UserData.UserPrincipalName; 'Recipient Type' = $UserData.RecipientTypeDetails; 'Archive Name' = $UserData.ArchiveName; 'Archive Mailbox Size' = $ArchiveMailboxSizeRounded; 'Archive Mailbox Size(Bytes)' = $ArchiveMailboxSizeBytes; 'Archive Quota' = $ArchiveQuotaSize; 'Archive Warning Quota'= $ArchiveWarningQuotaSize; 'Archive Status' = $UserData.ArchiveStatus; 'Archive State'= $UserData.ArchiveState; 'Auto Expanding Archive' = $AutoExpandArchive }
 $ExportResults = New-Object PSObject -Property $ExportResult
 $ExportResults | Select-Object 'DisplayName', 'Email Address', 'Recipient Type', 'Archive Name', 'Archive Mailbox Size', 'Archive Mailbox Size(Bytes)', 'Archive Quota','Archive Warning Quota','Archive Status','Archive State', 'Auto Expanding Archive' | Export-csv -path $ExportCSVFileName -NoType -Append
}

#Checking Exchange Online PowerShell Module Availability
$Exchange = (get-module ExchangeOnlineManagement -ListAvailable).Name
if($Exchange -eq $null) 
{
 Write-host "Important: Module ExchangeOnline is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
 $confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
 if($confirm -match "[yY]") 
 { 
  Write-host "Installing ExchangeOnlineManagement"
  Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
  Write-host "Required Module is installed in the machine Successfully"
 }
 else
 { 
  Write-host "Exiting. `nNote: Exchange module must be available in your system to run the script" 
  Exit 
 } 
}

#Storing credential in script for scheduling purpose/Passing credential as parameter
if(($UserName -ne "") -and ($Password -ne "")) 
{   
 $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force   
 $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword 
 Connect-ExchangeOnline -Credential $Credential -ShowProgress $false | Out-Null
}
else 
{
 Connect-ExchangeOnline
}
Write-Host "Exchange PowerShell Connected Successfully"
#End of Connecting Exchange Online

$ExportCSVFileName = ".\ArchiveMailboxSizeReport_$((Get-Date -format MMM-dd` hh-mm` tt).ToString()).csv"

Write-Host Generating report... `n
#filtering the conditions based on the User Input param
$WhereObjectCheck = $null
if($UserMBOnly.IsPresent -or $SharedMBOnly.IsPresent)
{
 $RecipientFilterVal = 'UserMailbox'
 if($SharedMBOnly.IsPresent) 
 {
  $RecipientFilterVal = 'SharedMailbox'
 }
 if($AutoExpandingArchiveEnabled.IsPresent) 
 {
  $WhereObjectCheck = { ($_.RecipientTypeDetails -eq $RecipientFilterVal) -and ($_.AutoExpandingArchiveEnabled -eq $true) }
 } 
 else 
 {
  $WhereObjectCheck = { $_.RecipientTypeDetails -eq $RecipientFilterVal }
 }
}
elseif($AutoExpandingArchiveEnabled.IsPresent)
{
    $WhereObjectCheck = { $_.AutoExpandingArchiveEnabled -eq $true }
}

$global:ReportSize =0

#Check for input file
if ([string]$MBIdentityFile -ne "") 
{ 
 CSVImport
}

#Generating all Archive Mailbox Size report
else 
{
 if ($null -eq $WhereObjectCheck) 
 {
  $AllValidMailboxes = Get-Mailbox -ResultSize Unlimited -Archive
 }
#Generating Specific Archive Mailbox Size report on checking Where condition
 else 
 {
  $AllValidMailboxes = Get-Mailbox -ResultSize Unlimited -Archive | Where-Object $WhereObjectCheck
 }
 foreach($UserData in $AllValidMailboxes) 
 {
  ExportOutput
 }
}

<#if($global:ReportSize -eq 0)
{
 Write-Host "There is no Archive Mailbox to return" 
}
else
{
 Write-Host "The output file contains $global:ReportSize mailboxes."`n
}
#>

#Open output file after execution
if((Test-Path -Path $ExportCSVFileName) -eq "True") 
{ 
 Write-Host "The output file contains $global:ReportSize mailboxes."`n
 Write-Host " The Output file available in:" -NoNewline -ForegroundColor Yellow
 Write-Host $ExportCSVFileName 
 $prompt = New-Object -ComObject wscript.shell    
 $userInput = $prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)    
 If ($userInput -eq 6)    
 {    
  Invoke-Item "$ExportCSVFileName"
 }  
}
else
{
 Write-Host "There is no Archive Mailbox to return" 
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n