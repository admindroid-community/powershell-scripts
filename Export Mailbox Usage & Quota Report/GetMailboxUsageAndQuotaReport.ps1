<#
=============================================================================================

Name         : Export Mailbox Quota Report Using PowerShell 
Version      : 1.0
website      : o365reports.com

-----------------
Script Highlights
-----------------
1. The script exports mailbox usage quota report indicating if the usage exceeds the set limits or not.
2. Exports mailboxes that exceed the ‘Issue Warning’ quota.
3. Retrieves mailboxes surpassing the ‘Prohibit Send’ quota.
4. Helps to identify mailboxes nearing the ‘Warning’ quota.
5. Automatically installs the MS Graph PowerShell module upon your confirmation when it is not available on your machine.
6. The script can be executed with MFA enabled account too.
7. Exports report results to CSV.
8. The script is scheduler friendly.
9. Also, you can execute this script with certificate-based authentication (CBA).

For detailed Script execution:  https://o365reports.com/2024/07/30/export-mailbox-quota-report-using-powershell/
============================================================================================
\#>

param (
    [string] $CertificateThumbPrint,
    [string] $ClientId,
    [string] $TenantId,
    [Switch] $OverIssueWarningQuota,
    [Switch] $OverProhibitSendQuota,
    [decimal] $NearingWarningQuota
)

#Check for MGgraph installation
$Module=Get-Module -Name Microsoft.Graph.Beta.Reports -ListAvailable
    if($Module.count -eq 0) 
    { 
        Write-Host Microsoft Graph PowerShell SDK is not available  -ForegroundColor yellow  
        $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
        if($Confirm -match "[yY]") 
        { 
            Write-Host "Installing Microsoft Graph Beta Reports powerShell module..."
            Install-Module Microsoft.Graph.Beta.Reports -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
        }
        else
        {
            Write-Host "Microsoft Graph Beta Reports PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph.Beta.Reports cmdlet." 
            Exit
        }
    }

Write-Host "Connecting to Microsoft.Graph."

#connect to Mggraph 
try{
    if(($ClientId -ne "") -and ($CertificateThumbPrint -ne "") -and ($TenantId -ne ""))
    {
        #connect mggraph using certificate based authentication 
        Connect-MgGraph -TenantId $TenantId -ClientID $ClientId -CertificateThumbprint $CertificateThumbPrint -ErrorAction Stop|Out-Null
    }
    else
    {
        Connect-MgGraph -scopes "Reports.Read.All" -ErrorAction Stop |Out-Null
    }
}
catch{
    Write-Host "Error occurred: $($_.Exception.Message )" -ForegroundColor Red
    Exit
}

Write-Host "Microsoft.Graph connected successfully." -ForegroundColor Green


Function byteConversion{
    param (
        [long]$bytes
    )
    if ($bytes -lt 1KB) {
        return "$bytes bytes"
    }
    elseif ($bytes -lt 1MB) {
        $KBValue = [math]::Round($bytes / 1KB, 2)
        return "$KBValue KB"
    }
    elseif ($bytes -lt 1GB) {
        $MBValue = [math]::Round($bytes / 1MB, 2)
        return "$MBValue MB"
    }
    else {
        $GBValue = [math]::Round($bytes / 1GB, 2)
        return "$GBValue GB"
    }
}

Function FindMailbox {
    Import-CSV $global:CSVReport | ForEach-Object {
        Write-Progress -Activity "Checking mailbox:" -status $_.'Display Name'
        $QuotaStatus = ""
        $StorageUsed = [long] $_.'Storage Used (Byte)'
        if($NearingWarningQuota){
            $IssueWarningQuota = [long] $_.'Issue Warning Quota (Byte)'
            $NearStorageValue = $IssueWarningQuota * ($NearingWarningQuota/100)
            if($StorageUsed -ge $NearStorageValue){
                $QuotaStatus = "StroageUsed Exceeds $NearingWarningQuota %"
            }                
        }
        else{
            if($global:Flag){
                $ProhibitSendReceiveQuota = [long] $_.'Prohibit Send/Receive Quota (Byte)'
                if($StorageUsed -eq $ProhibitSendReceiveQuota){
                    $QuotaStatus = "Reached Prohibit Send/Receive Quota"
                }
            }
            if($OverProhibitSendQuota -and ($QuotaStatus -eq "")){
                $Prohibitsendquota =[long] $_.'Prohibit Send Quota (Byte)'
                if($StorageUsed -ge $Prohibitsendquota){
                    $QuotaStatus = "Over Prohibit Send Quota"
                }
            }
            if($OverIssueWarningQuota -and ($QuotaStatus -eq "")){
                $Issuewarningquota =[long] $_.'Issue warning quota (Byte)'
                if($StorageUsed -ge $Issuewarningquota){
                    $QuotaStatus = "Over Issue Warning Quota"
                }
            }
            if(($QuotaStatus -eq "") -and $global:Flag){
                $QuotaStatus = "Within bounds"
            }
        }
        $AvailableStorage = [long] $_.'Prohibit Send/Receive Quota (Byte)' - $StorageUsed
        if($QuotaStatus -ne ""){ 
            #ExportResults
            $ExportResult = @{'User Principal Name' = $_.'User Principal Name'; 'Storage Used (Bytes)' = $StorageUsed; 'Storage Used' = byteConversion -bytes $storageUsed;'Available Storage' = byteConversion -bytes $AvailableStorage;'Item Count' = $_.'Item Count'; 'Quota Status' = $QuotaStatus; 'Issue Warning Quota (GB)' = [math]::Round($_.'Issue Warning Quota (Byte)' / 1GB, 2); 'Prohibit Send Quota (GB)' = [math]::Round( $_.'Prohibit Send Quota (Byte)' / 1GB, 2); 'Prohibit Send/Receive Quota (GB)' = [math]::Round($_.'Prohibit Send/Receive Quota (Byte)' / 1GB, 2); 'Display Name' = $_.'Display Name';'Recipient Type' = $_.'Recipient Type'; 'Is Mailbox Deleted' = $_.'Is Deleted'}
            $ExportResults = New-Object PSObject -Property $ExportResult
            $ExportResults | Select-object 'Display Name', 'User Principal Name', 'Storage Used', 'Storage Used (Bytes)','Available Storage', 'Issue Warning Quota (GB)', 'Prohibit Send Quota (GB)', 'Prohibit Send/Receive Quota (GB)','Quota Status', 'Recipient Type', 'Item Count', 'Is Mailbox Deleted' | Export-csv -path $global:CSV -NoType -Append -Force  
        }
    }
}

Function FetchReport{
    #Getting mailbox usage report
    try{
        Get-MgbetaReportMailboxUsageDetail -period "D180" -OutFile $global:CSVReport -ErrorAction stop
    }
    catch{
        Write-Host "Error occurred: $( $_.Exception.Message )" -ForegroundColor Red
        Disconnect-MgGraph out-Null
        Exit
    }
}

#................................Execution starts here...................................

if($OverIssueWarningQuota){
    $FileName = "Exceeds_IssueWarningQuota_"
}
elseif($OverProhibitSendQuota){
    $FileName = "Exceeds_ProhibitSendQuota_"
}
elseif($NearingWarningQuota){
    $FileName = "NearingTo_IssueWarningQuota_"
}
else{
    $FileName = "And_QuotaReport_"
}
$global:CSV = "MailboxStorageUsage_" + $FileName + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
$global:CSVReport = "MailboxStorageUsageDetailedReport_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"

$global:Flag = $false
if(!$NearingWarningQuota){
    if(!$OverIssueWarningQuota -and !$OverProhibitSendQuota){
        $global:Flag = $true
        $OverIssueWarningQuota = $true
        $OverProhibitSendQuota = $true
    }
}

Write-Host "Generating report..."
#To fetch mailbox usage report for reference
FetchReport
#To find mailbox quota status
FindMailbox
#Remove/delete the reference file
Remove-Item $global:CSVReport

$Location = Get-Location
if (((Test-Path -Path $global:CSV) -ne "True")) {     
    Write-Host "`nNo Mailboxes exceeds the storage usage for given criteria.`n" -ForegroundColor Green
}
else {
  
    Write-Host "`nThe output file $global:CSV is available in $Location" -ForegroundColor Yellow
        $prompt = New-Object -ComObject wscript.shell    
        $UserInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
        if ($UserInput -eq 6) {
            Invoke-Item "$global:CSV"
        }
    }
Disconnect-MgGraph | Out-Null
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n