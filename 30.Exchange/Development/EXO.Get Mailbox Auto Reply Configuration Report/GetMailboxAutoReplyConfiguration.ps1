<#
=============================================================================================
Name:           Export Office 365 mailbox users' OOF configuration status
Description:    This script exports Office 365 mailbox users' OOF configuration status to CSV
Version:        1.0
Website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Generates 5 different types of out of office set up status reports  
2. Automatically installs the Exchange Online module upon your confirmation when it is not available in your machine 
3. Lists aggregated result of users with enabled and scheduled auto-reply set up 
4. Retrieves users who have enabled auto-reply settings 
5. Generates users who configured scheduled out of office set up separately 
6. Lists currently unavailable users’ out of office configurations
7. You have an option to get the disabled mailboxes with active automatic reply setting 
8. Delivers Office 365 users’ upcoming out of office plans details 
9. Supports both MFA and Non-MFA accounts
10. Exports the report in CSV format
11. The script is scheduler-friendly. You can automate the report generation upon passing credentials as parameters.

For detailed Script execution: https://o365reports.com/2021/08/18/get-mailbox-automatic-reply-configuration-using-powershell
============================================================================================
#>

param (
    [string] $UserName = $null,
    [string] $Password = $null,
    [Switch] $Enabled,
    [Switch] $Scheduled,
    [Switch] $DisabledMailboxes,
    [Switch] $Today,
    [String] $ActiveOOFAfterDays
    
)

#Checks ExchangeOnline module availability and connects the module
Function ConnectToExchange {
    $Exchange = (get-module ExchangeOnlineManagement -ListAvailable).Name
    if ($Exchange -eq $null) {
        Write-host "Important: ExchangeOnline PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
        if ($confirm -match "[yY]") { 
            Write-host "Installing ExchangeOnlineManagement"
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
            Write-host "ExchangeOnline PowerShell module is installed in the machine successfully."`n
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
    Write-Host "ExchangeOnline PowerShell module is connected successfully"`n
    #End of Connecting Exchange Online
}

#This function checks the user choice and retrieves the OOF status
Function RetrieveOOFReport {
    #Checks the users with scheduled OOF setup
    if ($Scheduled.IsPresent) {
        $global:ExportCSVFileName = "OOFScheduledUsersReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
        Get-mailbox -ResultSize Unlimited | foreach-object {
            $CurrUser = $_
            $CurrOOFConfigData = Get-MailboxAutoReplyConfiguration -Identity ($CurrUser.PrimarySmtpAddress) | Where-object { $_.AutoReplyState -eq "Scheduled" }
            if ($null -ne $CurrOOFConfigData ) {
                PrepareOOFReport
                ExportScheduledOOF
            }
        }
    }
    #Checks the OOF status on and after user mentioned days 
    elseif ($ActiveOOFAfterDays -gt 0) {
        $global:ExportCSVFileName = "UpcomingOOFStatusReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
        $OOFStartDate = (Get-date).AddDays($ActiveOOFAfterDays).Date.ToString().split(" ") | Select -Index 0
        Get-mailbox -ResultSize Unlimited | foreach-object {
            $CurrUser = $_
            $CurrOOFConfigData = Get-MailboxAutoReplyConfiguration -Identity ($CurrUser.PrimarySmtpAddress) | Where-object { $_.AutoReplyState -ne "Disabled" }
            if ($null -ne $CurrOOFConfigData ) {
                $CurrOOFStartDate = $CurrOOFConfigData.StartTime.ToString().split(" ") | select -Index 0
                $CurrOOFEndDate = $CurrOOFConfigData.EndTime.ToString().split(" ") | Select -Index 0
                $ActiveOOFAfterFlag = "true"
                if($CurrOOFConfigData.AutoReplyState -eq "Enabled" -or ($OOFStartDate -ge $CurrOOFStartDate -and $OOFStartDate -le $CurrOOFEndDate)){
                    PrepareOOFReport
                    ExportAllActiveOOFSetup
                }
            }
        }
    }
    #Checks the OOF with enabled status
    elseif ($Enabled.IsPresent) {
        $global:ExportCSVFileName = "OOFEnabledUsersReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
        Get-mailbox -ResultSize Unlimited | foreach-object {
            $CurrUser = $_
            $CurrOOFConfigData = Get-MailboxAutoReplyConfiguration -Identity ($CurrUser.PrimarySmtpAddress) | Where-object { $_.AutoReplyState -eq "Enabled" } 
            if ($null -ne $CurrOOFConfigData ) {
                $EnabledFlag = 'true'
                PrepareOOFReport
                ExportEnabledOOF
            }
        }
    }
    #Checks whether OOF starting day is current day and process
    elseif ($Today.Ispresent) {
        $global:ExportCSVFileName = "OOFUsersTodayReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
        $CurrDate = (Get-Date).Date.ToString().split(" ") | Select -Index 0
        Get-mailbox -ResultSize Unlimited | foreach-object {
            $CurrUser = $_
            $CurrOOFConfigData = Get-MailboxAutoReplyConfiguration -Identity ($CurrUser.PrimarySmtpAddress) | Where-object { $_.AutoReplyState -ne "Disabled" }
            if ($null -ne $CurrOOFConfigData ) {
                $CurrOOFStartDate = $CurrOOFConfigData.StartTime.ToString().split(" ") | select -Index 0
                $CurrOOFEndDate = $CurrOOFConfigData.EndTime.ToString().split(" ") | Select -Index 0
                if ($CurrDate -ge $CurrOOFStartDate -and $CurrDate -le $CurrOOFEndDate) {
                    PrepareOOFReport
                    ExportAllActiveOOFSetup
                }
            }
        }
    }
    
    #Checks the disabled mailoxes OOF configuration
    elseif ($DisabledMailboxes.Ispresent) {
        $global:ExportCSVFileName = "DisabledAccountsOOFConfigurationReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
        Get-mailbox -ResultSize Unlimited | Where-Object { $_.AccountDisabled -eq $TRUE } | foreach-object {
            $CurrUser = $_
            $CurrOOFConfigData = Get-MailboxAutoReplyConfiguration -Identity ($CurrUser.PrimarySmtpAddress) | Where-object { $_.AutoReplyState -ne "Disabled" }
            if ($null -ne $CurrOOFConfigData ) {
                PrepareOOFReport
                ExportDisabledMailboxOOFSetup
            }
        }
    }   
    #Checks the all active OOF configuration
    else {
        $global:ExportCSVFileName = "OutofOfficeConfigurationReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
        Get-mailbox -ResultSize Unlimited | foreach-object {
            $CurrUser = $_
            $CurrOOFConfigData = Get-MailboxAutoReplyConfiguration -Identity ($CurrUser.PrimarySmtpAddress) | Where-object { $_.AutoReplyState -ne "Disabled" }
            if ($null -ne $CurrOOFConfigData ) {   
                PrepareOOFReport
                ExportAllActiveOOFSetup
            }
        }
    }
}

#Checks the Boolean values
Function GetPrintableValue($RawData) {
    if ($null -eq $RawData -or $RawData.Equals($false)) {
        return "No"
    }
    else {
        return "Yes"
    }
}

#Saves the users with OOF configuration
Function PrepareOOFReport {
    $global:ReportSize = $global:ReportSize + 1
    
    $EmailAddress = $CurrUser.PrimarySmtpAddress
    $AccountStatus = $CurrUser.AccountDisabled
    $MailboxOwner = $CurrOOFConfigData.MailboxOwnerId
    $OOFStatus = $CurrOOFConfigData.AutoReplyState
    $AutoCancelRequests = GetPrintableValue $CurrOOFConfigData.AutoDeclineFutureRequestsWhenOOF
    $CancelAllEvents = GetPrintableValue $CurrOOFConfigData.DeclineAllEventsForScheduledOOF
    $CancelScheduledEvents = GetPrintableValue $CurrOOFConfigData.DeclineEventsForScheduledOOF
    $CreateOOFEvent = GetPrintableValue $CurrOOFConfigData.CreateOOFEvent
    $ExternalAudience = $CurrOOFConfigData.ExternalAudience
    $StartTime = $CurrOOFConfigData.StartTime
    $EndTime = $CurrOOFConfigData.EndTime
    $Duration = $EndTime - $StartTime
    $TimeSpan = "$($Duration.Days.ToString('00'))d : $($Duration.Hours.ToString('00'))h : $($Duration.Minutes.ToString('00'))m";
    
    if ($CurrOOFConfigData.InternalMessage -ne "") {
        $InternalMessage = (($CurrOOFConfigData.InternalMessage) -replace '<.*?>', '').Trim()
    } 
    else { $InternalMessage = "-" }
    if ($CurrOOFConfigData.ExternalMessage -ne "") {
        $ExternalMessage = (($CurrOOFConfigData.ExternalMessage) -replace '<[^>]+>', '').Trim()
    }
    else { $ExternalMessage = "-" }
    
    if ($EnabledFlag -eq 'true') {
        $global:OOFDuration = 'OOF Duration'
    }
    else {
        $global:OOFDuration = 'OOF Duration (Days:Hours:Mins)'
    }
    if ($OOFStatus -eq 'Enabled') {
        $StartTime = "-"
        $EndTime = "-"
        $TimeSpan = 'Until auto-reply is disabled'
    }
                
    Write-Progress "Retrieving the OOF Status of the User: $MailboxOwner" "Processed Users Count: $global:ReportSize"

    #Save values with output column names 
    $ExportResult = @{ 

        'Email Address'            = $EmailAddress;
        'Disabled Account'         = $AccountStatus;
        'Mailbox Owner'            = $MailboxOwner;
        'Auto Reply State'         = $OOFStatus;
        'Start Time'               = $StartTime;
        'End Time'                 = $EndTime;
        'Decline Future Requests'  = $AutoCancelRequests;
        'Decline All Events'       = $CancelAllEvents;
        'Decline Scheduled Events' = $CancelScheduledEvents;
        'Create OOF Event'         = $CreateOOFEvent;
        'External Audience'        = $ExternalAudience;
        'Internal Message'         = $InternalMessage;
        'External Message'         = $ExternalMessage;
        $global:OOFDuration        = $TimeSpan 
    }

    $global:ExportResults = New-Object PSObject -Property $ExportResult
}

#Exports the users with OOF schedued configuration 
Function ExportScheduledOOF {
    $global:ExportResults | Select-object 'Mailbox Owner', 'Email Address', 'Start Time', 'End Time', $global:OOFDuration, 'Decline Future Requests', 'Decline All Events', 'Decline Scheduled Events', 'Create OOF Event', 'External Audience', 'Internal Message', 'External Message', 'Disabled Account' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
}

#Exports the users with OOF Enabled configuration
Function ExportEnabledOOF {
    $global:ExportResults | Select-object 'Mailbox Owner', 'Email Address', $global:OOFDuration, 'External Audience', 'Internal Message', 'External Message', 'Disabled Account' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
}

#Exports all the users with OOF configuration
Function ExportAllActiveOOFSetup {
    if($ActiveOOFAfterFlag = "true"){
        $global:ExportResults | Select-object 'Mailbox Owner', 'Email Address', 'Auto Reply State', 'Start Time', 'End Time', $global:OOFDuration, 'External Audience', 'Internal Message', 'External Message','Disabled Account' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
    }
    else{
        $global:ExportResults | Select-object 'Mailbox Owner', 'Email Address', 'Auto Reply State', 'Start Time', 'End Time', $global:OOFDuration, 'External Audience', 'Internal Message', 'External Message', 'Decline Future Requests', 'Decline All Events', 'Decline Scheduled Events', 'Create OOF Event', 'Disabled Account' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
    }
} 

Function ExportDisabledMailboxOOFSetup {
    $global:ExportResults | Select-object 'Mailbox Owner', 'Email Address', 'Auto Reply State', 'Start Time', 'End Time', $global:OOFDuration, 'Decline Future Requests', 'Decline All Events', 'Decline Scheduled Events', 'Create OOF Event', 'External Audience', 'Internal Message', 'External Message' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force 
}

#Execution starts here
ConnectToExchange
$global:ReportSize = 0
RetrieveOOFReport

#Validates the output file
if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") {     
    #Open file after code execution finishes
    Write-Host " The output file available in:"-NoNewline -ForegroundColor Yellow; Write-Host $global:ExportCSVFileName 
    Write-host `n"Exported $global:ReportSize records to CSV." `n
    Write-Host "Disconnected active ExchangeOnline session" `n
    Write-Host ~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
    Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n   
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