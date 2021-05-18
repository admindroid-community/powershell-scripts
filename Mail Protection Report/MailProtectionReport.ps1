
param(
    [string] $UserName = $null,
    [string] $Password = $null,
    [Switch] $SpamEmailsSent,
    [Switch] $SpamEmailsReceived,
    [Switch] $MalwareEmailsSent,
    [Switch] $MalwareEmailsReceived,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate
)

Function DateAndSwitchesValidation {
    $global:MaxStartDate = ((Get-Date).Date).AddDays(-10)
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
    if(-Not(($SpamEmailsSent.IsPresent)-or($SpamEmailsReceived.IsPresent)-or($MalwareEmailsSent.IsPresent)-or($MalwareEmailsReceived.IsPresent))){
        Write-Host "Exiting.`nNote: Choose one report to generate. Please try again" -ForegroundColor Red
        Exit
    }
    GetSpamMalwareData -StartDate $StartDate -EndDate $EndDate
}
Function GetSpamMalwareData {
    param (
        [DateTime]$StartDate,
        [DateTime]$EndDate
    )

    $EndDate = Get-Date $EndDate -Hour 23 -Minute 59 -Second 59
    
    ConnectToExchange
    $global:ExportedEmails = 0
    $global:Domain = "Recipient Domain"

    if ($SpamEmailsReceived.IsPresent) {
        $global:ExportCSVFileName = ".\SpamEmailsReceivedReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving spam emails received from $StartDate to $EndDate..."
        Get-MailDetailSpamReport -StartDate $StartDate -EndDate $EndDate -Direction Inbound -PageSize 5000 | ForEach-Object {
            $global:Domain = "Sender Domain"
            $CurrRecord = $_
            RetrieveEmailInfo
        }  
    }
    elseif ($MalwareEmailsReceived.IsPresent) {
        $global:ExportCSVFileName = ".\MalwareEmailsReceivedReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving malware emails received from $StartDate to $EndDate..."
        Get-MailDetailMalwareReport -StartDate $StartDate -EndDate $EndDate -Direction Inbound -PageSize 5000 | ForEach-Object {
            $global:Domain = "Sender Domain"
            $CurrRecord = $_
            RetrieveEmailInfo
        }   
    }
    elseif ($SpamEmailsSent.IsPresent) {
        $global:ExportCSVFileName = ".\SpamEmailsSentReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving spam emails sent from $StartDate to $EndDate..."
        Get-MailDetailSpamReport -StartDate $StartDate -EndDate $EndDate -Direction Outbound -PageSize 5000 | ForEach-Object {
            $CurrRecord = $_
            RetrieveEmailInfo
        }  
    }
    elseif ($MalwareEmailsSent.IsPresent) {
        $global:ExportCSVFileName = ".\MalwareEmailsSentReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
        Write-Host "Retrieving malware emails sent from $StartDate to $EndDate..."
        Get-MailDetailMalwareReport -StartDate $StartDate -EndDate $EndDate -Direction Outbound -PageSize 5000 | ForEach-Object {
            $CurrRecord = $_
            RetrieveEmailInfo
        }  
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
    Write-Host "ExchangeOnline PowerShell module is connected successfully"
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
    else{
        $Domain = $RecipientAddress.split("@") | Select-object -Index 1
    }
    ExportResults
}
Function ExportResults {
    $global:ExportedEmails = $global:ExportedEmails + 1
    $ExportResult = @{'Date' = $DateTime; 'Sender Address' = $SenderAddress; 'Recipient Address' = $RecipientAddress; 'Subject'= $Subject; 'Event Type' = $EventType; $global:Domain = $Domain}
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-Object 'Date', 'Sender Address', 'Recipient Address', 'Subject', 'Event Type',$global:Domain | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force  
}

DateAndSwitchesValidation


#Open output file after execution
if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") { 
    Write-Host "The output file available in $global:ExportCSVFileName" -ForegroundColor Green 
    Write-Host "The exported report has $global:ExportedEmails email details"
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