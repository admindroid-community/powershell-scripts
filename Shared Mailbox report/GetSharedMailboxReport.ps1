<#
=============================================================================================
Name:           Export all shared mailboxes in Office 365
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.The script uses modern authentication to connect to Exchange Online.    
2.The script can be executed with an MFA-enabled account too.    
3.Exports report results to CSV file.    
4.It helps to identify shared mailboxes with licenses separately.  
5.Helps to track email forwarding configured shared mailboxes. 
6.Automatically installs the EXO V2 module (if not installed already) upon your confirmation.   
7.The script is scheduler-friendly. I.e., Credentials can be passed as a parameter instead of getting interactively.

For detailed script execution:  https://o365reports.com/2022/07/13/get-shared-mailbox-in-office-365-using-powershell
============================================================================================
#>Param
(
    [Parameter(Mandatory = $false)]
    [switch]$LicensedOnly,
    [switch]$EmailForwardingEnabled,
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
Connect_Exo
$OutputCSV=".\SharedMailboxReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$Count=0
$OutputCount=0
$ExportResult=""   
$ExportResults=@()

#Retrieve all the shared mailboxes
Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox | foreach {
 $Print=$true
 $Count++
 $Name=$_.Name
 Write-Progress -Activity "`n     Processing $Name.."`n" Processed mailbox count: $Count"
 $PrimarySMTPAddress=$_.PrimarySMTPAddress
 $UPN=$_.UserPrincipalName
 $Alias=$_.Alias
 $ArchiveStatus=$_.ArchiveStatus
 $LitigationHold=$_.LitigationHoldEnabled
 $RetentionHold=$_.RetentionHoldEnabled
 $IsLicensed=$_.SkuAssigned
 $AuditEnabled=$_.AuditEnabled
 $HideFromAddressList=$_.HiddenFromAddressListsEnabled
 $ForwardingAddress=$_.ForwardingAddress
 $ForwardingSMTPAddress=$_.ForwardingSMTPAddress
 if($ForwardingAddress -eq $null)
 {
  $ForwardingAddress="-"
 }
 if($ForwardingSMTPAddress -eq $null)
 {
  $ForwardingSMTPAddress="-"
 }

 $MBSize=((Get-MailboxStatistics -Identity $UPN).totalItemSize.value)
 $MailboxSize=$MBSize.ToString().split("(") | Select-Object -Index 0
  If(($LicensedOnly.IsPresent) -and ($IsLicensed -eq $false))
 {
  $Print=$false
 }
 if(($EmailForwardingEnabled.IsPresent) -and (($ForwardingAddress -eq "-")-and ($ForwardingSMTPAddress -eq "-")))
 {
  $Print=$false
 }

   #Export result to csv
  if($Print -eq $true)
  {
   $OutputCount++
   $ExportResult=@{'Name'=$Name;'Primary SMTP Address'=$PrimarySMTPAddress;'Alias'=$Alias;'MB Size'=$MailboxSize;'Is Licensed'=$IsLicensed ;'Archive Status'=$ArchiveStatus;'Hide From Address List'=$HideFromAddressList;'Audit Enabled'=$AuditEnabled; 'Forwarding Address'=$ForwardingAddress;'Forwarding SMTP Address'=$ForwardingSMTPAddress;'Litigation Hold'=$LitigationHold;'Retention Hold'=$RetentionHold}
   $ExportResults= New-Object PSObject -Property $ExportResult  
   $ExportResults | Select-Object 'Name','Primary SMTP Address','Alias','MB Size','Is Licensed','Archive Status','Hide From Address List','Audit Enabled','Forwarding Address','Forwarding SMTP Address','Litigation Hold','Retention Hold' | Export-Csv -Path $OutputCSV -Notype -Append 
  }
}
 
#Open output file after execution
If($OutputCount -eq 0)
{
 Write-Host No records found
}
else
{
 Write-Host `nThe output file contains $OutputCount shared mailboxes.
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