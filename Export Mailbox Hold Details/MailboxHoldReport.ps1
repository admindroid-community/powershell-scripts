<#
=============================================================================================
Name:           Export Office 365 Email Holds Report
Description:    This script exports Office 365 Email Holds to CSV
Version:        1.0
website:        o365reports.com
Script by:      O365Reports Team
For detailed Script execution: http://o365reports.com/2021/06/28/export-office-365-mailbox-holds-report-using-powershell
============================================================================================
#>

param (
    [string] $UserName = $null,
    [string] $Password = $null,
    [Switch] $LitigationHoldsOnly,
    [Switch] $InPlaceHoldsOnly,
    [Switch] $RetentionHoldsOnly
)

Function GetBasicData {
    $global:ExportedMailbox = $global:ExportedMailbox + 1
    $global:MailboxName = $_.Name 
    $global:RecipientTypeDetails = $_.RecipientTypeDetails
    $global:UPN = $_.UserPrincipalName
}
Function RetrieveAllHolds {
    if ($LitigationHoldsOnly.IsPresent) {
        $global:ExportCSVFileName = "LitigationHoldsReport" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
            Get-mailbox -IncludeInactiveMailbox -ResultSize Unlimited | Where-Object { $_.LitigationHoldEnabled -eq $True } | foreach-object {
            $CurrLitigationHold = $_
            GetLitigationHoldsReport
        }
    }
    elseif ($InPlaceHoldsOnly.IsPresent) {
        $global:ExportCSVFileName = "InPlaceHoldsReport" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Get-mailbox -IncludeInactiveMailbox -ResultSize Unlimited | Where-Object { $_.InPlaceHolds -ne $Empty } | foreach-object {
            $CurrInPlaceHold = $_
            GetInPlaceHoldsReport
        }
    }
    elseif ($RetentionHoldsOnly.IsPresent) {
        $global:ExportCSVFileName = "RetentionHoldsReport" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Get-mailbox -IncludeInactiveMailbox -ResultSize Unlimited | Where-Object { $_.RetentionHoldEnabled -eq $True } | foreach-object {
            $CurrRetentionHold = $_
            GetRetentionHoldsReport
        }
    }
    else {
        $global:ExportCSVFileName = "AllActiveHoldsReport" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Get-mailbox -IncludeInactiveMailbox -ResultSize Unlimited | Where-Object { $_.LitigationHoldEnabled -eq $True -or $_.RetentionHoldEnabled -eq $True -or $_.InPlaceHolds -ne $Empty -or $_.ComplianceTagHoldApplied -eq $True } | foreach-object {
            $CurrMailbox = $_
            GetDefaultReport
        }
    }
}

Function GetInPlaceHoldType($HoldGuidList) {
    $HoldTypes = @()
    $InPlaceHoldCount = 0
    $HoldGuidList | ForEach-Object {
        $InPlaceHoldCount = $InPlaceHoldCount + 1
        if ($_ -match "UniH") {
            $HoldTypes += "eDiscovery Case"
        }
        elseif ($_ -match "^mbx") {
            $HoldTypes += "Specific Location Retention Policy"
        }
        elseif ($_ -match "^\-mbx") {
            $HoldTypes += "Mailbox Excluded Retention Policy"
        }
        elseif ($_ -match "skp") {
            $HoldTypes += "Retention Policy on Skype"
        }
        else {
            $HoldTypes += "In-Place Hold"
        }
    }
    
    return ($HoldTypes -join ", "), $InPlaceHoldCount
}

Function GetLitigationHoldsReport {
    GetBasicData
    $LitigationOwner = $CurrLitigationHold.LitigationHoldOwner
    if ($null -ne $CurrLitigationHold.LitigationHoldDate) {
        $LitigationHoldDate = ($CurrLitigationHold.LitigationHoldDate).ToString().Split(" ") | Select-Object -Index 0
    }
    $LitigationDuration = $CurrLitigationHold.LitigationHoldDuration
    if ($LitigationDuration -ne "Unlimited") {
        $LitigationDuration = ($LitigationDuration).split(".") | Select-Object -Index 0
    }

    Write-Progress "Retrieving the Litigation Hold Information for the User: $global:MailboxName" "Processed Mailboxes Count: $global:ExportedMailbox" 

    #ExportResult
    $ExportResult = @{'Mailbox Name' = $global:MailboxName; 'Mailbox Type' = $global:RecipientTypeDetails; 'UPN' = $global:UPN; 'Litigation Owner' = $LitigationOwner; 'Litigation Duration' = $LitigationDuration; 'Litigation Hold Date' = $LitigationHoldDate }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Mailbox Name', 'UPN', 'Mailbox Type',  'Litigation Owner', 'Litigation Duration', 'Litigation Hold Date' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force  
}

Function GetInPlaceHoldsReport {
    GetBasicData
    $InPlaceHoldInfo = GetInPlaceHoldType ($CurrInPlaceHold.InPlaceHolds)
    $InPlaceHoldType = $InPlaceHoldInfo[0]
    $NumberOfHolds = $InPlaceHoldInfo[1]

    Write-Progress "Retrieving the In-Place Hold Information for the User: $global:MailboxName" "Processed Mailboxes Count: $global:ExportedMailbox"

    #Export Results
    $ExportResult = @{'Mailbox Name' = $global:MailboxName; 'Mailbox Type' = $global:RecipientTypeDetails; 'UPN' = $global:UPN; 'Configured InPlace Hold Count' = $NumberOfHolds; 'InPlace Hold Type' = $InPlaceHoldType }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Mailbox Name', 'UPN', 'Mailbox Type',  'Configured InPlace Hold Count', 'InPlace Hold Type' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force 
}

Function GetRetentionHoldsReport {
    GetBasicData
    $RetentionPolicy = $CurrRetentionHold.RetentionPolicy
    $RetentionPolicyTag = ((Get-RetentionPolicy -Identity $RetentionPolicy).RetentionPolicyTagLinks) -join ","

    if (($CurrRetentionHold.StartDateForRetentionHold) -ne $Empy) {
        $StartDateForRetentionHold = ($CurrRetentionHold.StartDateForRetentionHold).ToString().split(" ") | Select-Object -Index 0
    }
    else {
        $StartDateForRetentionHold = "-"
    }
    if (($CurrRetentionHold.EndDateForRetentionHold) -ne $Empy) {
        $EndDateForRetentionHold = ($CurrRetentionHold.EndDateForRetentionHold).ToString().split(" ") | Select-Object -Index 0
    }
    else {
        $EndDateForRetentionHold = "-"   
    }

    Write-Progress "Retrieving the Retention Hold Information for the User: $global:MailboxName" "Processed Mailboxes Count: $global:ExportedMailbox"

    #ExportResult
    $ExportResult = @{'Mailbox Name' = $global:MailboxName; 'Mailbox Type' = $global:RecipientTypeDetails; 'UPN' = $global:UPN; 'Retention Policy Name' = $RetentionPolicy; 'Start Date for Retention Hold' = $StartDateForRetentionHold; 'End Date for Retention Hold' = $EndDateForRetentionHold; 'Retention Policy Tag' = $RetentionPolicyTag }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Mailbox Name', 'UPN', 'Mailbox Type',  'Retention Policy Name', 'Start Date for Retention Hold', 'End Date for Retention Hold', 'Retention Policy Tag' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force 
}

Function GetDefaultReport {
    GetBasicData
    $LitigationHold = $CurrMailbox.LitigationHoldEnabled
    $ComplianceTag = $CurrMailbox.ComplianceTagHoldApplied
    $RetentionHold = $CurrMailbox.RetentionHoldEnabled
    $ArchiveStatus = $CurrMailbox.ArchiveStatus
    $RetentionPolicy = $CurrMailbox.RetentionPolicy
    
    $LitigationDuration = $CurrMailbox.LitigationHoldDuration
    if ($LitigationDuration -ne "Unlimited") {
        $LitigationDuration = ($LitigationDuration).split(".") | Select-Object -Index 0
    }
    $InPlaceHold = $CurrMailbox.InPlaceHolds
    if ($InPlaceHold -ne $Empty) {
        $InPlaceHold = "Enabled"
    }
    else {
        $InPlaceHold = "Disabled"
    }

    Write-Progress "Retrieving All Active Holds Applied on the User: $global:MailboxName" "Processed Mailboxes Count: $global:ExportedMailbox"
            
    #ExportResult
    $ExportResult = @{'Mailbox Name' = $global:MailboxName; 'Mailbox Type' = $global:RecipientTypeDetails; 'UPN' = $global:UPN; 'Archive Status' = $ArchiveStatus; 'Litigation Hold Enabled' = $LitigationHold; 'Compliance Tag Hold Enabled' = $ComplianceTag; 'Retention Hold Enabled' = $RetentionHold; 'Litigation Duration' = $LitigationDuration; 'In-Place Hold Status' = $InPlaceHold; 'Retention Policy Name' = $RetentionPolicy }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Mailbox Name', 'UPN', 'Mailbox Type',  'Archive Status', 'Litigation Hold Enabled', 'In-Place Hold Status', 'Retention Hold Enabled', 'Compliance Tag Hold Enabled', 'Litigation Duration', 'Retention Policy Name' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force 

}
Function ConnectToExchange {
    $Exchange = (get-module ExchangeOnlineManagement -ListAvailable).Name
    if ($Exchange -eq $null) {
        Write-host "Important: ExchangeOnline PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
        if ($confirm -match "[yY]") { 
            Write-host "Installing ExchangeOnlineManagement"
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
            Write-host "ExchangeOnline PowerShell module is installed in the machine successfully."
        }
        elseif ($confirm -cnotmatch "[yY]" ) { 
            Write-host "Exiting. `nNote: ExchangeOnline PowerShell module must be available in your system to run the script." 
            Exit 
        }
    }
    #Storing credential in script for scheduling purpose/Passing credential as parameter
    if (($UserName -ne "") -and ($Password -ne "")) {   
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force   
        $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword 
        Connect-ExchangeOnline -Credential $Credential -ShowProgress $false | Out-Null
    }
    else {
        Connect-ExchangeOnline | Out-Null
    }
    Write-Host "ExchangeOnline PowerShell module is connected successfully"
    #End of Connecting Exchange Online
}

ConnectToExchange
$global:ExportedMailbox = 0
RetrieveAllHolds
  
if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") {     
    Write-Host "The output file available in $global:ExportCSVFileName" -ForegroundColor Green 
    write-host "There are $global:ExportedMailbox mailboxes records in the output file" 
    $prompt = New-Object -ComObject wscript.shell    
    $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
    If ($userInput -eq 6) {    
        Invoke-Item "$global:ExportCSVFileName"
    }  
}
else {
    Write-Host "No data found with the specified criteria"
}
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
Write-Host "Disconnected active ExchangeOnline session"
