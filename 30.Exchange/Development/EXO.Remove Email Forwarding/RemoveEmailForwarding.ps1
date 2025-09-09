<#
=============================================================================================
Name:           Remove email forwarding in Office 365
Version:        1.0
Website:        m365scripts.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.The script uses modern authentication to connect to Exchange Online. 
2.The script can be executed with MFA enabled account too.  
3.Exports the report result to a CSV file.  
4.Removes email forwarding configurations as well as disables the inbox rule with email forwarding. 
5.Removes forwarding from a specific user. 
6.Disables email forwarding for a list of users through input CSV. 
7.Automatically installs the EXO V2 module (if not installed already) upon your confirmation. 
8.Credentials are passed as parameters (scheduler-friendly). 

For detailed script execution:  https://m365scripts.com/exchange-online/remove-email-forwarding-in-office-365-using-powershell/
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$UserName = $NULL,
    [string]$Password = $NULL,
    [string]$Name = $NULL,
    [string]$CSV = $NULL
)

function WriteToLogFile ($message) {
    $message >> $logfile
}

Function Connect_Exo {
    #Check for EXO v2 module inatallation
    $Module = Get-Module ExchangeOnlineManagement -ListAvailable
    if ($Module.count -eq 0) { 
        Write-Host "Exchange Online PowerShell V2 module is not available"  -ForegroundColor yellow  
        $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No" 
        if ($Confirm -match "[yY]") { 
            Write-host "Installing Exchange Online PowerShell module"
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
        } 
        else { 
            Write-Host "EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet." 
            Exit
        }
    } 
    Write-Host "Connecting to Exchange Online..."
    #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
    if (($UserName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword
        Connect-ExchangeOnline -Credential $Credential
    }
    else {
        Connect-ExchangeOnline
    }
}

Function GetMailboxForwardingInfoAndRemoveForwarding {
    $MailboxInfo = Get-Mailbox $Name | Where-Object { $_.ForwardingSMTPAddress -ne $Empty -or $_.ForwardingAddress -ne $Empty }
    if ($MailboxInfo.count -ne 0) {
        $MailboxOwner = $MailboxInfo.Name
        Write-Progress -Activity "Processing Mailbox Forwarding for the User: $MailboxOwner" " "
        $global:ReportSize1 = $global:ReportSize1 + 1
        $DeliverToMailbox = $MailboxInfo.DeliverToMailboxandForward 
        if ($null -eq $MailboxInfo.ForwardingSMTPAddress) {
            $ForwardingSMTPAddress = "-"
        }
        else{
            $ForwardingSMTPAddress = (($MailboxInfo.ForwardingSMTPAddress).split(":") | Select -Index 1)
        }
        if ($null -eq $MailboxInfo.ForwardingAddress) {
            $ForwardTo = "-"
        }
        else{
            $ForwardTo = $MailboxInfo.ForwardingAddress
        }
        #ExportResults
        $ExportResult = @{'Mailbox Name' = $MailboxOwner; 'Forwarding SMTP Address' = $ForwardingSMTPAddress; 'Forward To' = $ForwardTo; 'Deliver To Mailbox and Forward' = $DeliverToMailbox }
        $ExportResults = New-Object PSObject -Property $ExportResult
        $ExportResults | Select-object 'Mailbox Name', 'Forwarding SMTP Address', 'Forward To', 'Deliver To Mailbox and Forward' | Export-csv -path $global:ExportCSVFileName1 -NoType -Append -Force

        #Remove Forwarding
        try {
            Set-Mailbox $Name -ForwardingAddress $NULL -ForwardingSmtpAddress $NULL -ErrorAction Stop -WarningAction SilentlyContinue
            if (($ForwardingSMTPAddress -ne "-")) {
                if (($ForwardTo -ne "-")) {
                    WriteToLogFile "Email forwarding successfully removed from $MailboxOwner - ForwardTo: $ForwardTo ForwardingSMTPAddress:$ForwardingSMTPAddress."
                }
                else {
                    WriteToLogFile "Email forwarding successfully removed from $MailboxOwner - ForwardingSMTPAddress:$ForwardingSMTPAddress."
                }
            }
            else {
                WriteToLogFile "Email forwarding successfully removed from $MailboxOwner - ForwardTo: $ForwardTo."
            }
        }
        catch {
            WriteToLogFile "Error occured while removing email forwarding configuration from $MailboxOwner."
        }
    }
}

Function GetInboxRulesInfoAndDisableForwarding {
    Get-InboxRule -Mailbox $Name | Where-Object { $_.ForwardAsAttachmentTo -ne $Empty -or $_.ForwardTo -ne $Empty -or $_.RedirectTo -ne $Empty } | ForEach-Object {
        Write-Progress "Processing the Inbox Rule for the User: $($_.MailboxOwnerId)" 
        $InboxRuleInfo = $_
        $MailboxOwner = $InboxRuleInfo.MailboxOwnerId
        $RuleName = $InboxRuleInfo.Name
        $Enable = $InboxRuleInfo.Enabled
        $RedirectTo = @()
        if ($null -ne $InboxRuleInfo.RedirectTo) {
            foreach($Temp in $InboxRuleInfo.RedirectTo){
                $RedirectTo = $RedirectTo + ((($Temp.split("[")) | Select-Object -Index 0)).Replace('"', '').Trim()
            }
        }
        else {
            $RedirectTo = "-"
        }
        $ForwardAsAttachment = @()
        if ($null -ne $InboxRuleInfo.ForwardAsAttachmentTo) {
            foreach($Temp in $InboxRuleInfo.ForwardAsAttachmentTo){
                $ForwardAsAttachment = $ForwardAsAttachment + ((($Temp.split("[")) | Select-Object -Index 0)).Replace('"', '').Trim()
            }
        }
        else {
            $ForwardAsAttachment = "-"
        }
        $ForwardTo = @()
        if ($null -ne $InboxRuleInfo.ForwardTo) {
            foreach($Temp in $InboxRuleInfo.ForwardTo){
                $ForwardTo = $ForwardTo + ((($Temp.split("[")) | Select-Object -Index 0)).Replace('"', '').Trim()
            }
        }
        $global:ReportSize2 = $global:ReportSize2 + 1
        $ExportResult = @{'Mailbox Name' = $MailboxOwner; 'Inbox Rule' = $RuleName; 'Forward As Attachment To' = $ForwardAsAttachment -join(","); 'Forward To' = $ForwardTo -join(","); 'Redirect To' = $RedirectTo -join(",") }
        $ExportResults = New-Object PSObject -Property $ExportResult
        $ExportResults | Select-object 'Mailbox Name', 'Inbox Rule', 'Forward To', 'Redirect To', 'Forward As Attachment To' | Export-csv -path $global:ExportCSVFileName2 -NoType -Append -Force
        #Disable Inbox Rule
        if ($Enable -eq 'True') {
            try {
                Disable-InboxRule -Identity $RuleName
                WriteToLogFile "The Inbox rule ($RuleName) present in $MailboxOwner mailbox is disabled."
            }
            catch { 
                WriteToLogFile "Error occured, while processing the Inbox rule ($RuleName) present in $MailboxOwner"
            }
        }
    }
}

Connect_Exo
$global:logfile = "RemoveForwardingLogFile_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".txt"
$global:ExportCSVFileName1 = "EmailForwardingConfigurationReport_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
$global:ExportCSVFileName2 = "InboxRulesWithForwarding_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
$global:ReportSize1 = 0
$global:ReportSize2 = 0
if (($Name -ne "")) {
    GetMailboxForwardingInfoAndRemoveForwarding 
    GetInboxRulesInfoAndDisableForwarding
}
elseif (($CSV -ne "")) {
    Import-Csv $CSV | ForEach-Object {
        $Name = $_.Name
        GetMailboxForwardingInfoAndRemoveForwarding 
        GetInboxRulesInfoAndDisableForwarding
    }
}
else {
    [string]$Name = Read-Host "Enter the user Name"
    GetMailboxForwardingInfoAndRemoveForwarding
    GetInboxRulesInfoAndDisableForwarding
}
if (((Test-Path -Path $global:ExportCSVFileName1) -ne "True") -and ((Test-Path -Path $global:ExportCSVFileName2) -ne "True") ) {     
    Write-Host "The given user mail is not forwarded to anyone." -ForegroundColor Green  
}
else {
    if ((Test-Path -Path $global:ExportCSVFileName1) -eq "True") {
        if ((Test-Path -Path $global:ExportCSVFileName2) -eq "True") {
            Write-Host `n "Following output files are generated and avaialble in the current directory:" -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCsv2"`n
            Write-Host "$global:ExportCSVFileName1 , $global:ExportCSVFileName2" -ForegroundColor Cyan
            Write-Host "The log file available in $global:logfile" -ForegroundColor Green
            $prompt = New-Object -ComObject wscript.shell    
            $userInput = $prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)    
            if ($userInput -eq 6) {    
                Invoke-Item "$global:ExportCSVFileName1"
                Invoke-Item "$global:ExportCSVFileName2"
                Invoke-Item "$global:logfile"
            }
        }
        else {
            Write-Host `n "The Output file available in: " -NoNewline -ForegroundColor Yellow; Write-Host "$global:ExportCSVFileName1"
			Write-Host `n "The log file available in: " -NoNewline -ForegroundColor Yellow; Write-Host "$global:logfile" 
            $prompt = New-Object -ComObject wscript.shell    
            $userInput = $prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)    
            if ($userInput -eq 6) {    
                Invoke-Item "$global:ExportCSVFileName1"
                Invoke-Item "$global:logfile"
            }
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
        }
    }
    else {
        Write-Host `n "The output file $global:ExportCSVFileName2 is available in  the current directory:" -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCsv2"`n
        if ((Test-Path -Path $global:logfile) -eq "True") {
            Write-Host "The log file available in $global:logfile" -ForegroundColor Green
            $prompt = New-Object -ComObject wscript.shell    
            $userInput = $prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)    
            if ($userInput -eq 6) {
                Invoke-Item "$global:ExportCSVFileName2"
                Invoke-Item "$global:logfile"
            }
        }
        else {
            $prompt = New-Object -ComObject wscript.shell    
            $userInput = $prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)    
            if ($userInput -eq 6) {
                Invoke-Item "$global:ExportCSVFileName2"
            }
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
        }
    }
}