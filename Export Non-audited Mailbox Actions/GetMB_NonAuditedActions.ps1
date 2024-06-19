<#
=============================================================================================
Name:           Export Exchange Online Non-audited mailbox Activities
Description:    This script exports non-audited mailbox activities to CSV file
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script uses modern authentication to connect to Exchange Online. 
2. The script can be executed with MFA enabled account. 
3. Exports the report result to a CSV file. 
4. Lists the non-audited mailbox actions for each logon type (Admin, Owner, Delegate). 
5. Helps to identify audit bypassed mailboxes. 
6. Automatically installs the EXO V2 module (if not installed already) upon your confirmation. 
7. Credentials are passed as parameters (scheduler-friendly), so worry not! i.e., credentials can be passed as parameters rather than being saved inside the script.

For detailed script execution: https://o365reports.com/2022/05/31/identify-non-audited-mailbox-activities-and-take-necessary-actions
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$UserName = $NULL,
    [string]$Password = $NULL,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint
)
$AuditAdmin = @("ApplyRecord", "Copy", "Create", "FolderBind", "HardDelete", "MailItemsAccessed", "Move", "MoveToDeletedItems", "RecordDelete", "Send", "SendAs", "SendOnBehalf", "SoftDelete", "Update", "UpdateCalendarDelegation", "UpdateFolderPermissions", "UpdateComplianceTag" , "UpdateInboxRules")
$AuditDelegate = @("ApplyRecord", "Create", "FolderBind", "HardDelete", "MailItemsAccessed", "Move", "MoveToDeletedItems", "RecordDelete", "SendAs", "SendOnBehalf", "SoftDelete", "Update", "UpdateFolderPermissions", "UpdateComplianceTag", "UpdateInboxRules")
$AuditOwner = @("ApplyRecord", "Create", "HardDelete", "MailboxLogin", "MailItemsAccessed", "Move", "MoveToDeletedItems", "RecordDelete", "Send", "SearchQueryInitiated", "SoftDelete", "Update", "UpdateCalendarDelegation", "UpdateFolderPermissions", "UpdateComplianceTag", "UpdateInboxRules")

function Connect_Exo {
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
$Location=Get-Location
$global:ExportCSVFileName = "$Location\Mailboxes_NonAuditingActions_Report_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
function MailboxNotAudited {
    $Audit_Check = Get-OrganizationConfig | Select AuditDisabled
    if ($Audit_Check.AuditDisabled -eq $true) {
        Write-Host "Auditing is disabled in your organization."
        Exit
    }
    else {
        $MBCount = 0
        Get-Mailbox -ResultSize Unlimited | ForEach-Object {
            $MBCount = $MBCount + 1
            $Identity = $_.UserPrincipalName
            $Name = $_.DisplayName
            Write-Progress -Activity "Processing Mailbox: $Name" -Status "Processed Mailbox Count: $MBCount" 
            $MBInfo = Get-Mailbox -Identity $Identity | Select-Object AuditOwner, AuditAdmin, AuditDelegate, DefaultAuditSet
            $Owner_ActionAudited = $MBInfo.AuditOwner
            $Admin_ActionAudited = $MBInfo.AuditAdmin
            $Delegate_ActionAudited = $MBInfo.AuditDelegate
            $DefaultAuditSet = $MBInfo.DefaultAuditSet
            $Owner_ActionNotAudited = $AuditOwner | Where-Object { $_ -notin $Owner_ActionAudited }
            $Admin_ActionNotAudited = $AuditAdmin | Where-Object { $_ -notin $Admin_ActionAudited }
            $Delegate_ActionNotAudited = $AuditDelegate | Where-Object { $_ -notin $Delegate_ActionAudited }
            $AuditByPassEnabled = Get-MailboxAuditBypassAssociation -Identity $Identity | Select-Object AuditByPassEnabled
            if ($Owner_ActionNotAudited.count -eq 0) {
                $Owner_ActionNotAudited = "-"
            }
            if ($Admin_ActionNotAudited.count -eq 0) {
                $Admin_ActionNotAudited = "-"
            }
            if ($Delegate_ActionNotAudited.count -eq 0) {
                $Delegate_ActionNotAudited = "-"
            }
            if($DefaultAuditSet.count -eq 0){
                $DefaultAuditSet = "-"
            }
            $ExportResult = @{'Display Name' = $Name; 'Logon type with Default Audit Set' = $DefaultAuditSet -join(","); 'Audit By Pass Enabled' = $AuditByPassEnabled.AuditByPassEnabled; 'Owner' = $Owner_ActionNotAudited -join (","); 'Admin' = $Admin_ActionNotAudited -join (","); 'Delegate' = $Delegate_ActionNotAudited -join (",") }
            $ExportResults = New-Object PSObject -Property $ExportResult
            $ExportResults | Select-object 'Display Name', 'Audit By Pass Enabled', 'Logon type with Default Audit Set', 'Owner', 'Admin', 'Delegate' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
        }
    }
}

Connect_Exo
MailboxNotAudited
if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") {     
    Write-Host "Mailboxes and disabled auditing actions are exported"`n
    Write-Host " The report available in:" -NoNewline -ForegroundColor Yellow; Write-Host $global:ExportCSVFileName `n
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
    Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    $prompt = New-Object -ComObject wscript.shell    
    $userInput = $prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)    
    if ($userInput -eq 6) {    
        Invoke-Item "$global:ExportCSVFileName"
    }  
}
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
