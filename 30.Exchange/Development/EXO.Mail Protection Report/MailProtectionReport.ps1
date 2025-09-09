<#
=============================================================================================
Name:           Export Office 365 Spam, Malware and phish Report using PowerShell 
Description:    This script exports Office 365 spam, malware and phish report  to CSV
Version:        2.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~

1.Generates 9 different email protection reports.
2.Automatically installs the Exchange Online PowerShell module upon your confirmation when it is not available in the system.
3.Supports both MFA and Non-MFA accounts.
4.Specify date ranges to generate reports for custom periods. 
5.Supports filters to retrieve sent and received spams. 
6.Allows you to filter sent and received malwares. 
7.Tracks sent and received phishing emails. 
8.Facilitates the separation of internal spam, malware, and phishing emails. 
9.Exports the report to CSV.   
10.Scheduler-friendly. You can automate the report generation upon passing credentials as parameters.  
 
For detailed script execution: https://o365reports.com/2021/05/18/export-office-365-spam-and-malware-report-using-powershell
============================================================================================
#>

param(
    [string] $UserName = $null,
    [string] $Password = $null,
    [Switch] $SpamEmailsSent,
    [Switch] $SpamEmailsReceived,
    [Switch] $MalwareEmailsSent,
    [Switch] $MalwareEmailsReceived,
    [Switch] $PhishEmailsSent,
    [Switch] $PhishEmailsReceived,
    [Switch] $IntraorgSpamMails,
    [Switch] $IntraorgMalwareMails,
    [Switch] $IntraorgPhishMails,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate
)

Function DateAndSwitchesValidation {
    $global:MaxStartDate = ((Get-Date).Date).AddDays(-30)
    if (($StartDate -eq $Null) -and ($EndDate -eq $Null)) {
        $StartDate = $global:MaxStartDate
        $EndDate = (Get-date).Date
    }
    elseif (($StartDate -eq $null) -or ($EndDate -eq $null)) {
            Write-Host "Exiting.`nNote: Both start and end date values are mandatory. Please try again." -ForegroundColor Red
            Exit
        }
    elseif ($StartDate -lt $global:MaxStartDate) {
            Write-Host "Exiting.`nNote: You can retrieve data from $global:MaxStartDate onwards. Please try again." -ForegroundColor Red
            Exit
        }
    else {
         $StartDate = [DateTime]$StartDate
         $EndDate = [DateTime]$EndDate 
        }
    if(-Not(($SpamEmailsSent.IsPresent)-or($SpamEmailsReceived.IsPresent)-or($MalwareEmailsSent.IsPresent)-or($MalwareEmailsReceived.IsPresent)-or($PhishEmailsSent.IsPresent)-or($PhishEmailsReceived.IsPresent)-or($IntraorgSpamMails.IsPresent)-or($IntraorgMalwareMails.IsPresent)-or($IntraorgPhishMails.IsPresent))){
        Write-Host "Exiting.`nNote: Choose one report to generate. Please try again" -ForegroundColor Red
        Exit
    }
    GetSpamMalwarePhishData -StartDate $StartDate -EndDate $EndDate
}
Function GetSpamMalwarePhishData {
    param (
        [DateTime]$StartDate,
        [DateTime]$EndDate
    )

    $EndDate = Get-Date $EndDate -Hour 23 -Minute 59 -Second 59
    
    ConnectToExchange
    $global:ExportedEmails = 0
    $global:Domain = "Recipient Domain"
    $SpamEventTypes = "URL malicious reputation", "Advanced filter", "General filter", "Mixed analysis detection", "Fingerprint matching", "Domain reputation", "Bulk", "IP reputation"
    $PhishEventTypes = "URL malicious reputation", "Advanced filter", "General filter", "Spoof intra-org", "Spoof external domain", "Spoof DMARC", "Impersonation brand", "Mixed analysis detection", "File reputation", "Fingerprint matching", "URL detonation reputation", "URL detonation", "Impersonation user", "Impersonation domain", "Mailbox intelligence impersonation", "File detonation", "File detonation reputation", "Campaign"
    $MalwareEventTypes = "File detonation", "File detonation reputation", "File reputation", "Anti-malware engine", "URL malicious reputation", "URL detonation", "URL detonation reputation", "Campaign"

    if ($SpamEmailsReceived.IsPresent) {
        $global:ExportCSVFileName = ".\SpamEmailsReceivedReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving spam emails received from $StartDate to $EndDate..."`n
        Get-MailDetailATPReport -StartDate $StartDate -EndDate $EndDate -Direction Inbound -PageSize 5000 -EventType $SpamEventTypes | Where-Object { $_.VerdictSource -like "Spam"} | ForEach-Object {
            $global:Domain = "Sender Domain"
            $CurrRecord = $_
            RetrieveEmailInfo
        }
        OpenOutputFile
    }
    if ($MalwareEmailsReceived.IsPresent) {
        $global:ExportCSVFileName = ".\MalwareEmailsReceivedReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving malware emails received from $StartDate to $EndDate..."`n
        Get-MailDetailATPReport -StartDate $StartDate -EndDate $EndDate -Direction Inbound -PageSize 5000 -EventType $MalwareEventTypes | Where-Object { $_.VerdictSource -like "Malware"} | ForEach-Object {
            $global:Domain = "Sender Domain"
            $CurrRecord = $_
            RetrieveEmailInfo
        }
        OpenOutputFile
    }
    if ($SpamEmailsSent.IsPresent) {
        $global:ExportCSVFileName = ".\SpamEmailsSentReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving spam emails sent from $StartDate to $EndDate..."`n
        Get-MailDetailATPReport -StartDate $StartDate -EndDate $EndDate -Direction Outbound -PageSize 5000 -EventType $SpamEventTypes | Where-Object { $_.VerdictSource -like "Spam"} | ForEach-Object {
            $CurrRecord = $_
            RetrieveEmailInfo
        }
        OpenOutputFile
    }
    if ($MalwareEmailsSent.IsPresent) {
        $global:ExportCSVFileName = ".\MalwareEmailsSentReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving malware emails sent from $StartDate to $EndDate..."`n
        Get-MailDetailATPReport -StartDate $StartDate -EndDate $EndDate -Direction Outbound -PageSize 5000 -EventType $MalwareEventTypes | Where-Object { $_.VerdictSource -like "Malware"} | ForEach-Object {
            $CurrRecord = $_
            RetrieveEmailInfo
        }
        OpenOutputFile
    }
    if ($PhishEmailsReceived.IsPresent) {
        $global:ExportCSVFileName = ".\PhishEmailsReceivedReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving phish emails received from $StartDate to $EndDate..."`n
        Get-MailDetailATPReport -StartDate $StartDate -EndDate $EndDate -Direction Inbound -PageSize 5000 -EventType $PhishEventTypes | Where-Object { $_.VerdictSource -like "Phish"} | ForEach-Object {
            $global:Domain = "Sender Domain"
            $CurrRecord = $_
            RetrieveEmailInfo
        }
        OpenOutputFile
    }
    if ($PhishEmailsSent.IsPresent) {
        $global:ExportCSVFileName = ".\PhishEmailsSentReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving phish emails sent from $StartDate to $EndDate..."`n
        Get-MailDetailATPReport -StartDate $StartDate -EndDate $EndDate -Direction Outbound -PageSize 5000 -EventType $PhishEventTypes | Where-Object { $_.VerdictSource -like "Phish"} | ForEach-Object {
            $CurrRecord = $_
            RetrieveEmailInfo
        }
        OpenOutputFile
    }
    if ($IntraorgSpamMails.IsPresent) {
        $global:ExportCSVFileName = ".\IntraorgSpamMailsReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving internal spam emails from $StartDate to $EndDate..."`n
        Get-MailDetailATPReport -StartDate $StartDate -EndDate $EndDate -Direction IntraOrg -PageSize 5000 -EventType $SpamEventTypes | Where-Object { $_.VerdictSource -like "Spam"} | ForEach-Object {
            $CurrRecord = $_
            RetrieveEmailInfo
        }
        OpenOutputFile
    }
    if ($IntraorgMalwareMails.IsPresent) {
        $global:ExportCSVFileName = ".\IntraorgMalwareMailsReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving internal malware emails from $StartDate to $EndDate..."`n
        Get-MailDetailATPReport -StartDate $StartDate -EndDate $EndDate -Direction IntraOrg -PageSize 5000 -EventType $MalwareEventTypes | Where-Object { $_.VerdictSource -like "Malware"} | ForEach-Object {
            $CurrRecord = $_
            RetrieveEmailInfo
        }
        OpenOutputFile
    }
    if ($IntraorgPhishMails.IsPresent) {
        $global:ExportCSVFileName = ".\IntraorgPhishMailsReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving internal phish emails from $StartDate to $EndDate..."`n
        Get-MailDetailATPReport -StartDate $StartDate -EndDate $EndDate -Direction IntraOrg -PageSize 5000 -EventType $PhishEventTypes | Where-Object { $_.VerdictSource -like "Phish"} | ForEach-Object {
            $CurrRecord = $_
            RetrieveEmailInfo
        }
        OpenOutputFile
    }
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
    Write-Host "ExchangeOnline PowerShell module is connected successfully"`n
    #End of Connecting Exchange Online
}
Function RetrieveEmailInfo {
    $Date = $CurrRecord.Date.ToShortDateString()
    $DateTime = $CurrRecord.Date
    Write-Progress -Activity "Retrieving mail data for $Date"
    $SenderAddress = $CurrRecord.SenderAddress
    $RecipientAddress = $CurrRecord.RecipientAddress
    $Subject = $CurrRecord.Subject
    $EventType = $CurrRecord.EventType
    if($CurrRecord.Direction -eq 'Inbound'){
        $Domain = $SenderAddress.split("@") | Select-object -Index 1
    }
    elseif($CurrRecord.Direction -eq 'Outbound'){
        $Domain = $RecipientAddress.split("@") | Select-object -Index 1
    }
    ExportResults
}
Function ExportResults {
    $global:ExportedEmails = $global:ExportedEmails + 1
    $ExportResult = @{'Date' = $DateTime; 'Sender Address' = $SenderAddress; 'Recipient Address' = $RecipientAddress; 'Subject'= $Subject; 'Event Type' = $EventType; $global:Domain = $Domain}
    $ExportResults = New-Object PSObject -Property $ExportResult
    if(($CurrRecord.Direction -eq 'Inbound')-or($CurrRecord.Direction -eq 'Outbound')){
        $ExportResults | Select-Object 'Date', 'Sender Address', 'Recipient Address', 'Subject', 'Event Type',$global:Domain | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force  
    }
    else{
        $ExportResults | Select-Object 'Date', 'Sender Address', 'Recipient Address', 'Subject', 'Event Type' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force  
    }
}
#Open output file after execution
Function OpenOutputFile{
    if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") { 
        Write-Host " The Output file available in:" -NoNewline -ForegroundColor Yellow; Write-Host $global:ExportCSVFileName
        Write-Host `n"The exported report has $global:ExportedEmails email details"
        $prompt = New-Object -ComObject wscript.shell    
        $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
        If ($userInput -eq 6) {    
            Invoke-Item "$global:ExportCSVFileName"
        }  
    }
    else {
        Write-Host "No data found with the specified criteria"
    }
}

DateAndSwitchesValidation

Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
Write-Host `n"Disconnected active ExchangeOnline session"
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n