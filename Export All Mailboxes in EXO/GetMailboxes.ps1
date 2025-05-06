<#
=============================================================================================
Name:           Export All Mailboxes in Microsoft 365 
Version:        1.0
Website:        o365reports.com


Script Highlights:  
~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs Exchange Online PowerShell module (if not installed already) upon your confirmation.    
2. Exports all mailboxes in the organization. 
3. Allows exporting mailboxes that matches the selected filter. 
	- User Mailboxes 
	- Shared Mailboxes 
	- Room Mailboxes 
	- Equipment Mailboxes   
4. Export the report result to CSV. 
5. The script can be executed with MFA enabled account too.  
6. The script supports certificate-based authentication (CBA) too. 
7. The script is schedular-friendly.   

For detailed Script execution: https://o365reports.com/2025/05/06/export-all-mailboxes-in-microsoft-365-using-powershell


============================================================================================
#>


Param
(
    [Parameter(Mandatory = $false)]
    [switch]$UserMailboxesOnly,
    [switch]$SharedMailboxesOnly,
    [Switch]$RoomMailboxesOnly,
    [Switch]$EquipmentMailboxesOnly,
    [string]$UserName,
    [string]$Password,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint
)
Function Connect_Exo
{
 #Check for EXO module installation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell  module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
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
  Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
 }
 elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
   Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
 }
 else
 {
  Connect-ExchangeOnline -ShowBanner:$false
 }
}
Connect_Exo

$OutputCSV="$PSScriptRoot\EXOMailbox_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$Count=0

#Filtering based on mailbox type
if($RoomMailboxesOnly.IsPresent)
{
 $RecipientType="RoomMailbox"
}
elseif($UserMailboxesOnly.IsPresent)
{
 $RecipientType="UserMailbox"
}
elseif($SharedMailboxesOnly.IsPresent)
{
 $RecipientType="SharedMailbox"
}
elseif($EquipmentMailboxesOnly.IsPresent)
{
 $RecipientType="EquipmentMailbox"
}
else
{
 $RecipientType="DiscoveryMailbox,EquipmentMailbox,GroupMailbox,LegacyMailbox,LinkedMailbox,LinkedRoomMailbox,RoomMailbox,SchedulingMailbox,SharedMailbox,TeamMailbox,UserMailbox"
}

#Retrieving mailbox details
Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $RecipientType | foreach {
 $DisplayName=$_.DisplayName
 $Count++
  Write-Progress -Activity "`n  Processed mailbox count: $Count .."`n" Currently processing: $DisplayName"
 $UPN=$_.UserPrincipalName
 $Alias=$_.Alias
 $MailboxType=$_.RecipientTypeDetails
 $PrimarySMTPAddress=$_.PrimarySMTPAddress
 $EmailAddresses= ($_.EmailAddresses | Where-Object {$_ -clike "smtp:*"} | ForEach-Object {$_ -replace "smtp:",""}) -join ","
 If($EmailAddresses -eq "")
 {
  $EmailAddresses="-"
 }

 $Results = [PSCustomObject]@{
                            "Mailbox Name"             = $DisplayName
                            "UPN"          = $UPN
                            "Alias" =$Alias
                            "Mailbox Type"         = $MailboxType
                            "Primary SMTP Address"      = $PrimarySMTPAddress
                            "Other Email Addresses" =$EmailAddresses
                           
 }
 $Results | Export-CSV  -path $OutputCSV -NoTypeInformation -Append  -Force
}

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n
  
If($Count -eq 0)
{
 Write-Host No mailboxes found
}
else
{
 Write-Host `nThe output file contains $Count mailboxes.
 Write-Host `n" The Output file available in: "  -NoNewline -ForegroundColor Yellow; Write-Host $OutputCSV 
 $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
 If ($UserInput -eq 6)   
 {   
  Invoke-Item "$OutputCSV"   
 } 
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue