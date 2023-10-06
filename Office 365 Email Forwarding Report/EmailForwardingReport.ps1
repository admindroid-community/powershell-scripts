<#
=============================================================================================
Name:           Export Office 365 Email Forwarding Report using PowerShell 
Description:    This script exports Office 365 email forwarding report  to CSV format
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.Generates 3 different email forwarding rules reports.  
2.Automatically installs the Exchange Online module upon your confirmation when it is not available in your machine. 
3.Shows mailboxes in which email forwarding configured through ‘Forwarding SMTP Address’ and ‘Forward To’. 
4.Lists all inbox rules that forward email to others’ mailbox. 
5.Identifies transport rule that redirects emails to mailboxes 
6.Supports both MFA and Non-MFA accounts.    
7.Exports the report in CSV format.  
8.The script is scheduler-friendly. You can automate the report generation upon passing credentials as parameters. 

For detailed script execution: https://o365reports.com/2021/06/09/export-office-365-email-forwarding-report-using-powershell/
============================================================================================
#>

param(
    [string] $UserName = $null,
    [string] $Password = $null,
    [Switch] $InboxRules,
    [Switch] $MailFlowRules
)


Function GetPrintableValue($RawData) {
    if (($null -eq $RawData) -or ($RawData.Equals(""))) {
        return "-";
    }
    else {
        $StringVal = $RawData | Out-String
        return $StringVal;
    }
}

Function GetAllMailForwardingRules {
    Write-host `n"Preparing the Email Forwarding Report..."
    if($InboxRules.IsPresent) {
        $global:ExportCSVFileName = "InboxRulesWithEmailForwarding_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Get-Mailbox -ResultSize Unlimited | ForEach-Object { 
            Write-Progress "Processing the Inbox Rule for the User: $($_.Id)" " "
            Get-InboxRule -Mailbox $_.PrimarySmtpAddress | Where-Object { $_.ForwardAsAttachmentTo -ne $Empty -or $_.ForwardTo -ne $Empty -or $_.RedirectTo -ne $Empty} | ForEach-Object {
                $CurrUserRule = $_
                GetInboxRulesInfo
            }
        }
    }
    Elseif ($MailFlowRules.IsPresent) {
        $global:ExportCSVFileName = "TransportRulesWithEmailForwarding_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Get-TransportRule -ResultSize Unlimited | Where-Object { $_.RedirectMessageTo -ne $Empty } | ForEach-Object {
            Write-Progress -Activity "Processing the Transport Rule: $($_.Name)" " "
            $CurrEmailFlowRule = $_
            GetMailFlowRulesInfo
        }
    } 
    else{
        $global:ExportCSVFileName = "EmailForwardingReport_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Get-Mailbox -ResultSize Unlimited | Where-Object { $_.ForwardingSMTPAddress -ne $Empty -or $_.ForwardingAddress -ne $Empty} | ForEach-Object {
            Write-Progress -Activity "Processing Mailbox Forwarding Rules for the User: $($_.Id)" " "
            $CurrEmailSetUp = $_
            GetMailboxForwardingInfo
        }
    }
}


Function GetMailboxForwardingInfo {
    $global:ReportSize = $global:ReportSize + 1
    $MailboxOwner = $CurrEmailSetUp.PrimarySMTPAddress
    $DeliverToMailbox = $CurrEmailSetUp.DeliverToMailboxandForward 
    if ($null -ne $CurrEmailSetUp.ForwardingSMTPAddress) {
        $CurrEmailSetUp.ForwardingSMTPAddress = GetPrintableValue (($CurrEmailSetUp.ForwardingSMTPAddress).split(":") | Select -Index 1)
    }
    $ForwardingSMTPAddress = GetPrintableValue $CurrEmailSetUp.ForwardingSMTPAddress
    if ($null -ne $CurrEmailSetUp.ForwardingAddress){
        $CurrEmailSetUp.ForwardingAddress = GetPrintableValue ($CurrEmailSetUp.ForwardingAddress)
    }
    $ForwardTo = GetPrintableValue $CurrEmailSetUp.ForwardingAddress
    
    #ExportResults
    $ExportResult = @{'Mailbox Name' = $MailboxOwner; 'Forwarding SMTP Address' = $ForwardingSMTPAddress;'Forward To' =$ForwardTo; 'Deliver To Mailbox and Forward' = $DeliverToMailbox}
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Mailbox Name', 'Forwarding SMTP Address','Forward To','Deliver To Mailbox and Forward' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force 
}

Function GetInboxRulesInfo {
    $global:ReportSize = $global:ReportSize + 1
    $MailboxOwner = $CurrUserRule.MailboxOwnerId
    $RuleName = $CurrUserRule.Name
    $Enable = $CurrUserRule.Enabled
    $StopProcessingRules = $CurrUserRule.StopProcessingRules
    if ($null -ne $CurrUserRule.RedirectTo) {
        $CurrUserRule.RedirectTo = GetPrintableValue (($CurrUserRule.RedirectTo).split("[") | Select-Object -Index 0).Replace('"', '').Trim()
    }
    $RedirectTo = GetPrintableValue $CurrUserRule.RedirectTo
    if ($null -ne $CurrUserRule.ForwardAsAttachmentTo) {
        $CurrUserRule.ForwardAsAttachmentTo = GetPrintableValue (($CurrUserRule.ForwardAsAttachmentTo).split("[") | Select-Object -Index 0).Replace('"', '').Trim()
    }
    $ForwardAsAttachment = GetPrintableValue $CurrUserRule.ForwardAsAttachmentTo
    if ($null -ne $CurrUserRule.ForwardTo) {
        $CurrUserRule.ForwardTo = GetPrintableValue (($CurrUserRule.ForwardTo).split("[") | Select-Object -Index 0).Replace('"', '').Trim()
    }
    $ForwardTo = GetPrintableValue $CurrUserRule.ForwardTo
    
    #ExportResults
    $ExportResult = @{'Mailbox Name' = $MailboxOwner; 'Inbox Rule' = $RuleName; 'Rule Status' = $Enable; 'Forward As Attachment To' = $ForwardAsAttachment; 'Forward To' = $ForwardTo; 'Stop Processing Rules' = $StopProcessingRules; 'Redirect To' = $RedirectTo }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Mailbox Name', 'Inbox Rule', 'Forward To', 'Redirect To', 'Forward As Attachment To','Stop Processing Rules', 'Rule Status' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force 
}

Function GetMailFlowRulesInfo {
    $global:ReportSize = $global:ReportSize + 1
    $RuleName = $CurrEmailFlowRule.Name
    $State = $CurrEmailFlowRule.State
    $Mode = $CurrEmailFlowRule.Mode
    $Priority = $CurrEmailFlowRule.Priority
    $StopProcessingRules = $CurrEmailFlowRule.StopRuleProcessing
    if ($null -ne $CurrEmailFlowRule.RedirectMessageTo) {
        $CurrEmailFlowRule.RedirectMessageTo = GetPrintableValue ($CurrEmailFlowRule.RedirectMessageTo).Replace('{}', '').Trim()
    }
    $RedirectTo = $CurrEmailFlowRule.RedirectMessageTo
    
    #ExportResults
    $ExportResult = @{'Mail Flow Rule Name' = $RuleName; 'State' = $State; 'Mode' = $Mode; 'Priority' = $Priority; 'Redirect To' = $RedirectTo; 'Stop Processing Rule' = $StopProcessingRules}
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Mail Flow Rule Name','Redirect To', 'Stop Processing Rule','State', 'Mode', 'Priority' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force 
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
$global:ReportSize = 0
GetAllMailForwardingRules
Write-Progress -Activity "--" -Completed

if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") {     
    Write-Host `n" The output file available in:" -NoNewline -ForegroundColor Yellow
    Write-Host	.\$global:ExportCSVFileName `n
    Write-Host "The exported report has $global:ReportSize email forwarding configurations"
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
    Write-Host `n"Disconnected active ExchangeOnline session"
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  
