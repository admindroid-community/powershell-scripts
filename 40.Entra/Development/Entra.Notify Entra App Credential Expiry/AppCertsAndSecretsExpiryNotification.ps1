<#
=============================================================================================
Name:           Send Microsoft Entra App Credentials Expiry Notifications  
Version:        1.1
Website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Sends app credential expiry notifications to specific users. 
2. Sends notifications for expiring certificates alone, client secrets alone, or both.
3. Exports a list of apps with expiring credentials within the specified days in CSV format.
4. Allows sending emails on behalf of other users.
5. Automatically install the Microsoft Graph PowerShell module (if not installed already) upon your confirmation.
6. The script can be executed with an MFA-enabled account too.
7. It can be executed with certificate-based authentication (CBA) too. 
8. The script is scheduler-friendly.

Change Log
~~~~~~~~~~
   V1.0 (Apr 29, 2025) - File created
   V1.1 (Jun 14, 2025) - Minor code improvements.



For detailed Script execution: https://o365reports.com/2025/04/29/send-entra-app-credential-expiry-notifications
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $True)]
    [int]$SoonToExpireInDays,
    [Parameter(Mandatory = $True)]
    [string]$Recipients,
    [string]$FromAddress,
    [Switch]$ClientSecretsOnly,
    [Switch]$CertificatesOnly,
    [Switch]$StoreReportLocally,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

$CSVFilePath ="$(Get-Location)\AppCertsAndSecretsExpiryNotificationSummary_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 

# Check if Microsoft Graph module is installed
$MsGraphModule = Get-Module Microsoft.Graph -ListAvailable
if ($MsGraphModule -eq $null) {
    Write-Host "`nImportant: Microsoft Graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
    $confirm = Read-Host "Are you sure you want to install Microsoft Graph module? [Y] Yes [N] No"
    if ($confirm -match "[yY]") {
        Write-Host "Installing Microsoft Graph module..."
        Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber
        Write-Host "Microsoft Graph module is installed in the machine successfully" -ForegroundColor Magenta 
    } else {
        Write-Host "Exiting. `nNote: Microsoft Graph module must be available in your system to run the script" -ForegroundColor Red
        Exit
    }
} 

Write-Host "`nConnecting to Microsoft Graph..."
    
if (($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne "")) {
    # Use certificate-based authentication if TenantId, ClientId, and CertificateThumbprint are provided
    Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
} else {
    # Use delegated permissions (Scopes) if credentials are not provided
    Connect-MgGraph -Scopes "Application.Read.All", "Mail.Send.Shared", "User.Read.All" -NoWelcome 
}

# Verify connection
if ((Get-MgContext) -ne $null) {
    if ((Get-MgContext).Account -ne $null) {
        $LoggedInAccount = (Get-MgContext).Account
        if([string]::IsNullOrEmpty($FromAddress)) {
            $FromAddress = $LoggedInAccount
        }
        Write-Host "Connected to Microsoft Graph PowerShell using account: $($LoggedInAccount)"
    }
    else {
        Write-Host "Connected to Microsoft Graph PowerShell using certificate-based authentication."
        if ([string]::IsNullOrEmpty($FromAddress)) {
            Write-Host "`nError: FromAddress is required when using certificate-based authentication." -ForegroundColor Red
            Exit
        }
    }
} else {
    Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
    Exit
}


# Function to Send Email
function SendEmail {
    $EmailAddresses = ($Recipients -split ",").Trim()
    $toRecipients = @()
    foreach ($Email in $EmailAddresses) {
        $toRecipients += @{
            emailAddress = @{
                address = $Email
            }
        }
    }

    $Script:TableContent += "</table>"
    $TableStyle = "<style>
        table { width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; }
        th, td { border: 1px solid black; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        </style>"
    
    $MailContent = "$($TableStyle)
        <p>Hello Admin,</p>
        <p>These application credentials are soon to expire:</p>
        $($Script:TableContent)
        <p>To prevent authentication failures and service disruptions, please renew the expiring secret or certificate via the <a href='https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade/quickStartType~/null/sourceType/Microsoft_AAD_IAM' target='_blank'>App registrations</a> in Microsoft Entra admin center.</p>  
        <p>If you have any questions, feel free to contact IT support.</p>  
        <p>Best regards,<br> IT Admin Team</p>"

    $params = @{
        message = @{
            subject = "Entra App Registration Credentials Expiry Notification"
            body = @{
                contentType = "HTML"
                content = $MailContent 
            }
            toRecipients = $toRecipients
        }
    }

    Send-MgUserMail -UserId $FromAddress -BodyParameter $params 
}


$ExportResult = $null   
$AppCount = 0
$Script:ProcessedCount = 0
$RequiredProperties=@('DisplayName','AppId','Id','KeyCredentials','PasswordCredentials','CreatedDateTime','SigninAudience')

if(($CertificatesOnly.IsPresent) -or ($ClientSecretsOnly.IsPresent) -or ($SoonToExpireInDays -ne "")) {
    $SwitchPresent=$True
}
else {
    $SwitchPresent=$false
}

# Create an HTML table with data
$Script:TableContent = "<table>"
$Script:TableContent += "<tr><th>App Name</th><th>App Creation Time</th><th>Credential Type</th><th>Credential Name</th><th>Creation Time</th><th>Expiry Date</th><th>Friendly Expiry Date</th></tr>"


Get-MgApplication -All -Property $RequiredProperties | ForEach-Object {
    $AppCount++
    $AppName=$_.DisplayName
    Write-Progress -Activity "`n     Processed App registration: $AppCount - $AppName"
    $AppId=$_.Id
    $Secrets=$_.PasswordCredentials
    $Certificates=$_.KeyCredentials
    $AppCreationDate=$_.CreatedDateTime
    $SigninAudience=$_.SignInAudience
    $AppOwners=(Get-MgApplicationOwner -ApplicationId $AppId).AdditionalProperties.userPrincipalName
    $Owners=$AppOwners -join ","
    
    if($owners -eq "") { $Owners="-" }
        
    #Process through Secret keys
    if(!($CertificatesOnly.IsPresent) -or ($SwitchPresent -eq $false)) {
        foreach($Secret in $Secrets) {
            $CredentialType="Client Secret"
            $DisplayName=$Secret.DisplayName
            $Id=$Secret.KeyId
            $CreatedTime=$Secret.StartDateTime
            $ExpiryDate=$Secret.EndDateTime
            $ExpiryStatusCalculation=(New-TimeSpan -Start (Get-Date).Date -End $ExpiryDate).Days
            $FriendlyExpiryTime="Expires in $ExpiryStatusCalculation days"

            if (($ExpiryStatusCalculation -gt 0) -and ($ExpiryStatusCalculation -le $SoonToExpireInDays)) { 
                $Script:ProcessedCount++
                if ($StoreReportLocally.IsPresent) {
                    $ExportResult = [PSCustomObject]@{'App Name'=$AppName;'App Owners'=$Owners;'App Creation Time'=$AppCreationDate;'Credential Type'=$CredentialType;'Credential Name'=$DisplayName;'Credential Id'=$Id;'Creation Time'=$CreatedTime;'Expiry Date'=$ExpiryDate;'Days to Expiry'=$ExpiryStatusCalculation;'Friendly Expiry Date'=$FriendlyExpiryTime;'App Id'=$AppId}
                    $ExportResult | Export-Csv -Path $CSVFilePath -Notype -Append
                }
                $Script:TableContent += "<tr><td>$($AppName)</td><td>$($AppCreationDate)</td><td>$($CredentialType)</td><td>$($DisplayName)</td><td>$($CreatedTime)</td><td>$($ExpiryDate)</td><td>$($FriendlyExpiryTime)</td></tr>"
             }
        }
    }

    #Process through Certificates
    if(!($ClientSecretsOnly.IsPresent) -or ($SwitchPresent -eq $false)){
        foreach ($Certificate in $Certificates) {
            $CredentialType="Certificate"
            $DisplayName=$Certificate.DisplayName
            $Id=$Certificate.KeyId
            $CreatedTime=$Certificate.StartDateTime
            $ExpiryDate=$Certificate.EndDateTime
            $ExpiryStatusCalculation=(New-TimeSpan -Start (Get-Date).Date -End $ExpiryDate).Days
            $FriendlyExpiryTime="Expires in $ExpiryStatusCalculation days"

            if (($ExpiryStatusCalculation -gt 0) -and ($ExpiryStatusCalculation -le $SoonToExpireInDays)) { 
                $Script:ProcessedCount++
                if ($StoreReportLocally.IsPresent) {
                    $ExportResult = [PSCustomObject]@{'App Name'=$AppName;'App Owners'=$Owners;'App Creation Time'=$AppCreationDate;'Credential Type'=$CredentialType;'Credential Name'=$DisplayName;'Credential Id'=$Id;'Creation Time'=$CreatedTime;'Expiry Date'=$ExpiryDate;'Days to Expiry'=$ExpiryStatusCalculation;'Friendly Expiry Date'=$FriendlyExpiryTime;'App Id'=$AppId}
                    $ExportResult | Export-Csv -Path $CSVFilePath -Notype -Append
                }
                $Script:TableContent += "<tr><td>$($AppName)</td><td>$($AppCreationDate)</td><td>$($CredentialType)</td><td>$($DisplayName)</td><td>$($CreatedTime)</td><td>$($ExpiryDate)</td><td>$($FriendlyExpiryTime)</td></tr>"
             }
        }
    }
}



if ($Script:ProcessedCount -ne 0) {
    SendEmail
    Write-Host `n"$Script:ProcessedCount app(s) credential is about to expire in $SoonToExpireInDays days." -ForegroundColor Yellow
    Write-Host `n"Email has been sent successfully." -ForegroundColor Yellow
    
    if ((Test-Path -Path $CSVFilePath) -eq "True") {
        Write-Host `n"The output file saved in: " -NoNewline -ForegroundColor Yellow
        Write-Host $CSVFilePath
    }
} else{
    Write-Host `n"No app(s) found with secrets or certificates expiring in the next $SoonToExpireInDays days." -ForegroundColor Yellow
}


# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1900+ Microsoft 365 reports. ~~" -ForegroundColor Green