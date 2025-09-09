<#
=============================================================================================
Name:           Trace Inbound External Emails in Exchange Online
Version:        1.1
Website:        o365reports.com


Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Exports emails received from external domains into a CSV file.
2. Supports exporting emails within a specified 90-day time range.
3. Lists emails received from a specific external domain.
4. Finds emails received from a specific external user.
5. Audit emails received by a specific user from external domains.
6. Allows filtering the mail flow report based on mail status.
7. The script automatically verifies and installs the Exchange Online PowerShell module (if not installed already) upon your confirmation.
8. The script can be executed with an MFA-enabled account.
9. It can be executed with Certificate-based Authentication (CBA) too.
10. The script is scheduler-friendly.
 
For detailed Script execution: https://o365reports.com/2025/06/03/trace-emails-received-from-external-domains-in-exchange-online

Change Log
~~~~~~~~~~

    V1.0 (Jun 03, 2025) - File created
    V1.1 (Jun 09, 2025) - Extended message tracing period from 30 to 90 days.
============================================================================================
#>

param(
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserName,
    [string]$Password,
    [string]$SenderAddress = "",
    [string]$RecipientAddress = "",
    [ValidateSet("Delivered", "Failed", "Gettingstatus", "FilteredAsSpam","Quarantined")]
    [string]$MailStatus,
    [string]$ExternalDomainName,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint

)

#--------------------------------------------------- Date Time Checking ---------------------------------------------------
$MaxStartDate = ((Get-Date).AddDays(-90)).Date

if (-not $StartDate -and -not $EndDate) {
    $EndDate = (Get-Date).Date
    $StartDate = $MaxStartDate
}

try {
    if ($StartDate) { $StartDate = [DateTime]$StartDate }
    if ($EndDate) { $EndDate = [DateTime]$EndDate }

    if ($StartDate -lt $MaxStartDate) {
        Write-Host "Error: MessageTrace can only be retrieved for the past 90 days. Select a date after $MaxStartDate" -ForegroundColor Red
        Exit
    }
    if ($EndDate -lt $StartDate) {
        Write-Host "Error: End date must be later than start date." -ForegroundColor Red
        Exit
    }
}
catch {
    Write-Host "Error: Invalid date format. Please enter a valid date." -ForegroundColor Red
    Exit
}

$CurrentStart = $StartDate
$CurrentEnd = $EndDate.AddDays(1).AddSeconds(-1)

#-------------------------------------------------- Function: Connect Exchange Online --------------------------------------
Function Connect_Exo {
    
																														  

    $installedModule = Get-Module ExchangeOnlineManagement -ListAvailable | Where-Object { $_.Version -ge [version]"3.0" }

    if (-not $installedModule) {
        Write-Host "Exchange Online PowerShell module is not available or unsupported version is available" -ForegroundColor Yellow
        $confirm = Read-Host "Do you want to install the latest module? [Y] Yes [N] No"
        if ($confirm -match "[yY]") {
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
        }
        else {
            Write-Host "EXO module is required. Please install it manually using 'Install-Module ExchangeOnlineManagement'."
            exit
        }
    }


    Write-Host "Connecting to Exchange Online..."
    
    if ($UserName -and $Password) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential = New-Object System.Management.Automation.PSCredential ($UserName, $SecuredPassword)
        Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
    }

    elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
   {
    Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
    }


    else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
}


#-------------------------------------------------- Function: Export To CSV -----------------------------------------------
Function Export-ToCSV {
    param (
        [Parameter(Mandatory = $true)]
        $Results
    )

    $Results | 
    Select-Object MessageTraceID, 
    @{Name = "Received Time"; Expression = { ($_.Received).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } }, 
    @{Name = "Sender Address"; Expression = { $_.SenderAddress } }, 
    @{Name = "Recipient Address"; Expression = { $_.RecipientAddress } }, 
    Subject, Status,
    @{Name = "Sender Domain"; Expression = { ($_.SenderAddress -split "@")[1] } },  
    @{Name = "Recipient Domain"; Expression = { ($_.RecipientAddress -split "@")[1] } },
    @{Name = "Sender IP"; Expression = { $_.FromIP } }, #FromIP
    @{Name = "Receipient IP"; Expression = { $_.ToIP } }, #TOIP
    @{Name = "Mail Size(KB)"; Expression = { [math]::Round($_.Size / 1KB, 2) } } | 
    Export-Csv -Path $OutputCSV -Append -Force -NoTypeInformation
}

#-------------------------------------------------- Function: Filter Data --------------------------------------------------
Function Filterdata {
   
    $FilteredResults = @()


    foreach ($queryResult in $queryResults) {

        $senderDomain = ($queryResult.SenderAddress -split "@")[-1]
        $recipientDomain = ($queryResult.RecipientAddress -split "@")[-1]

        if (
        ($InternalDomains -contains $recipientDomain) -and
        (-not ($InternalDomains -contains $senderDomain)) -and
        ([string]::IsNullOrEmpty($SenderAddress) -or $queryResult.SenderAddress -eq $SenderAddress) -and
        ([string]::IsNullOrEmpty($RecipientAddress) -or $queryResult.RecipientAddress -eq $RecipientAddress) -and
        ([string]::IsNullOrEmpty($ExternalDomainName) -or $ExternalDomainName -eq $SenderDomain) -and
        ([string]::IsNullOrEmpty($MailStatus) -or $queryResult.Status -eq $MailStatus)
        ) {
            $FilteredResults += $queryResult
        }
    }

    if ($FilteredResults.Count -gt 0) {
        $Global:FilteredCount += $FilteredResults.Count
        Export-ToCSV -Results $FilteredResults
    }

}

#-------------------------------------------------- Initialization ---------------------------------------------------------
Connect_Exo
$batchSize = 5000
$InternalDomains = (Get-AcceptedDomain).DomainName   # Gett all Internal Domains
$Location = Get-Location
$Timestamp = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")
$OutputCSV = "$Location\MailsReceivedFromExternalDomain_Report_$Timestamp.csv"
$queryResults = $null
$ProcessedCount = 0
$Global:FilteredCount = 0

#-------------------------------------------------- Core Processing Loop ---------------------------------------------------
while ($CurrentEnd -ge $CurrentStart) {
    $IntervalStartDate = $CurrentEnd.AddDays(-10)

    if ($IntervalStartDate -lt $CurrentStart) { $IntervalStartDate = $CurrentStart }   # To check currentstartdate not go beyond Actualstartdate 
    if ($IntervalStartDate -ge $CurrentEnd) { break } # To Terminate if DateInterval reaches enddate
    


    try {
        $queryResults = Get-MessageTraceV2 -StartDate $IntervalStartDate -EndDate $CurrentEnd -ResultSize $batchSize -ErrorAction Stop
        $ProcessedCount += $queryResults.Count

        if ($queryResults.Count -eq 0) {
            # If the current Quary has No result  change date to next Quary ,Skip the current Quary
            $CurrentEnd = $IntervalStartDate
            continue
        }

        Filterdata
    }
    catch {
        Write-Host "Error fetching message trace data: $_" -ForegroundColor Red
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Exit
    }

    Write-Progress -Activity "Retrieving Mail Record from $StartDate to $EndDate" -Status "Processed count: $ProcessedCount"

    ## #To Fetch Subsequant Data in the current quary

    while ($queryResults.Count -eq $batchSize) {
        $lastMessage = $queryResults[-1]
        $LastEndDate = $lastMessage.Received.ToString("O")
        $StartingRecipientAddress = $lastMessage.RecipientAddress

        try {
            $queryResults = Get-MessageTraceV2 -StartDate $IntervalStartDate -EndDate $LastEndDate -StartingRecipientAddress $StartingRecipientAddress -ResultSize $batchSize -ErrorAction Stop
            $ProcessedCount += $queryResults.Count
         
            Write-Progress -Activity "Retrieving Mail Record from $StartDate to $EndDate" -Status "Processed count: $ProcessedCount"
           
            # To break the loop if $quaryresult is zero
            if ($queryResults.Count -eq 0) { break }
            Filterdata
        }
        catch {
            Write-Host "Error fetching additional message trace data: $_" -ForegroundColor Red
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Exit
        }
    }

   

    $CurrentEnd = $IntervalStartDate   # Adjusted the End date Last IntervalStartdate
}

#-------------------------------------------------- Final Output -----------------------------------------------------------
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green																														  
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1900+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n

If ($FilteredCount -eq 0) {
    Write-Host "No records found"
}
else {
    Write-Host "`nThe output file contains $FilteredCount Mail records."
    if (Test-Path -Path $OutputCSV) {
        Write-Host "`nThe Output file is available at: " -NoNewline -ForegroundColor Yellow
        Write-Host $OutputCSV
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open the output file?", 0, "Open Output File", 4)
        If ($UserInput -eq 6) {
            Invoke-Item "$OutputCSV"
        }
    }
}

Disconnect-ExchangeOnline -Confirm:$false
