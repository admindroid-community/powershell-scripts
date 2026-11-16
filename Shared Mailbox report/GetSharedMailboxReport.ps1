<#
=============================================================================================
Name:           Export all shared mailboxes in Office 365
Version:        2.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.The script uses modern authentication to connect to Exchange Online.    
2.The script can be executed with an MFA-enabled account too.    
3.Exports report results to CSV file.    
4.It helps to identify shared mailboxes with licenses separately.  
5.Helps to track email forwarding configured shared mailboxes. 
6.Automatically installs the EXO module (if not installed already) upon your confirmation.   
7.Supports Certificate-based Authentication (CBA) and is scheduler-friendly.

Change Log:
~~~~~~~~~~
    V1.0 (Jul 13, 2022) - File created
    V2.0 (Nov 16, 2026) - Fixed the export stopping at 1000 shared mailboxes, migrated to EXO cmdlets, and replaced username/password authentication with CBA

For detailed script execution:  https://o365reports.com/2022/07/13/get-shared-mailbox-in-office-365-using-powershell
============================================================================================
#>Param
(
    [Parameter(Mandatory = $false)]
    [switch]$LicensedOnly,
    [switch]$EmailForwardingEnabled,
    [switch]$FullAccessOnly,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

# Module check

Function Connect-ExchangeOnlineSession {
    $ExchangeOnlineModule = Get-Module ExchangeOnlineManagement -ListAvailable
    if ($null -eq $ExchangeOnlineModule) {
        Write-Host "`nImportant: Exchange Online module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."
        $Confirm = Read-Host "Are you sure you want to install Exchange Online module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-Host "Installing Exchange Online module..."
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -AllowClobber
            Write-Host "Exchange Online module is installed in the machine successfully." -ForegroundColor Magenta
        }
        else {
            Write-Host "Exiting.`nNote: Exchange Online module must be available in your system to run the script." -ForegroundColor Red
            Exit
        }
    }

    Import-Module ExchangeOnlineManagement

    Write-Host "`nConnecting to Exchange Online..."
    if ($UseCBA) {
        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint `
            -Organization $Organization -ShowBanner:$false -ErrorAction Stop
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }
    Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
}

# CBA needs all three application parameters.
$CBAParameters = @($Organization, $ClientId, $CertificateThumbprint) |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
$UseCBA = $CBAParameters.Count -eq 3

if ($CBAParameters.Count -gt 0 -and -not $UseCBA) {
    Write-Host "Certificate-based authentication requires -Organization, -ClientId, and -CertificateThumbprint together." -ForegroundColor Red
    return
}

Connect-ExchangeOnlineSession
$OutputCSV=".\SharedMailboxReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$Count=0
$OutputCount=0
$ExportResult=""   
$ExportResults=@()

#Retrieve all the shared mailboxes before processing
$SharedMailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox -Properties DisplayName,ArchiveStatus,LitigationHoldEnabled,RetentionHoldEnabled,SkuAssigned,AuditEnabled,HiddenFromAddressListsEnabled,ForwardingAddress,ForwardingSmtpAddress
foreach ($Mailbox in $SharedMailboxes) {
 $Print=$true
 $Count++
 $Name=$Mailbox.DisplayName
 Write-Progress -Activity "`n     Processing $Name.."`n" Processed mailbox count: $Count"
 $PrimarySMTPAddress=$Mailbox.PrimarySMTPAddress
 $UPN=$Mailbox.UserPrincipalName
 $Alias=$Mailbox.Alias
 $ArchiveStatus=$Mailbox.ArchiveStatus
 $LitigationHold=$Mailbox.LitigationHoldEnabled
 $RetentionHold=$Mailbox.RetentionHoldEnabled
 $IsLicensed=$Mailbox.SkuAssigned
 $AuditEnabled=$Mailbox.AuditEnabled
 $HideFromAddressList=$Mailbox.HiddenFromAddressListsEnabled
 $ForwardingAddress=$Mailbox.ForwardingAddress
 $ForwardingSMTPAddress=$Mailbox.ForwardingSMTPAddress
 if($ForwardingAddress -eq $null)
 {
  $ForwardingAddress="-"
 }
 if($ForwardingSMTPAddress -eq $null)
 {
  $ForwardingSMTPAddress="-"
 }

 $MBSize=((Get-EXOMailboxStatistics -Identity $UPN).TotalItemSize)
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
 
Write-Host "`n~~ Script prepared by AdminDroid Community ~~`n" -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green
Write-Host "admindroid.com" -NoNewline -ForegroundColor Yellow
Write-Host " to get access to 3500+ Microsoft 365 reports and 450+ management actions. ~~`n`n" -ForegroundColor Green

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
 }
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue