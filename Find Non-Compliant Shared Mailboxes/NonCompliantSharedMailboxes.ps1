<#
=============================================================================================
Name:           Identify Non-Compliant Shared Mailboxes in Microsoft 365  
Version:        1.0
Website:        o365reports.com


Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Generates all non-compliant shared mailboxes in Microsoft 365.  
2. Exports report results to CSV file. 
3. The script automatically verifies and installs the MS Graph PowerShell SDK and Exchange Online PowerShell modules (if they are not already installed) upon your confirmation. 
4. The script can be executed with an MFA-enabled account too. 
5. The script supports Certificate-based authentication (CBA). 
6. The script is scheduler-friendly.   

For detailed Script execution: https://o365reports.com/2024/12/10/identify-non-compliant-shared-mailboxes-in-microsoft-365/


============================================================================================
#>Param
(
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [string]$TenantId,
    [string]$Organization,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password
)

# Function to connect to Microsoft Graph
function Connect_ToMgGraph {
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
        Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All"  -NoWelcome 
    }

    # Verify connection
    if ((Get-MgContext) -ne $null) {
        Write-Host "Connected to Microsoft Graph successfully." 
    } else {
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
}

# Function to connect to Exchange Online
function Connect-Exo {
    try {
        # Check for Exchange Online module installation
        $ExchangeModule = Get-Module ExchangeOnlineManagement -ListAvailable
        if ($ExchangeModule.count -eq 0) {
            Write-Host "ExchangeOnline module is not available" -ForegroundColor Yellow
            $confirm = Read-Host "Do you want to Install ExchangeOnline module? [Y] Yes  [N] No"
            if ($confirm -match "[Yy]") {
                Write-Host "Installing ExchangeOnline module ..."
                Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
                Import-Module ExchangeOnlineManagement
            } else {
                Write-Host "ExchangeOnline Module is required. To Install ExchangeOnline module use 'Install-Module ExchangeOnlineManagement' cmdlet."
                Exit
            }
        }

        # Connect Exchange Online module
        Write-Host "`nConnecting to Exchange Online module ..." 
        if (($UserName -ne "") -and ($Password -ne "")) {
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword
            Connect-ExchangeOnline -Credential $Credential
        } elseif ($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "") {
            Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
        } else {
            Connect-ExchangeOnline -ShowBanner:$false
        }

        Write-Host "Connected to Exchange Online successfully." 
    } catch {
        Write-Host "Error connecting to Exchange Online: $_" -ForegroundColor Red
        Exit
    }
}

# Main Execution
Connect_ToMgGraph
Connect-Exo

$ProgressIndex = 0
$OutputCSV = "$(Get-Location)\NonCompliant_Shared_Mailboxes_$((Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')).csv"
$NonCompliantCount = 0


$ExchangeOnlineServicePlans = @(
    "efb87545-963c-4e0d-99df-69c6916d9eb0", # EXCHANGE_S_ENTERPRISE(EXO Plan2)
    "9aaf7827-d63c-4b61-89c3-182f06f82e5c"  # EXCHANGE_S_STANDARD  (EXO Plan1)
)

# Retrieve all shared mailboxes
Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox | ForEach-Object {
    $MailboxName = $_.DisplayName
    $ProgressIndex++
    Write-Progress -Activity "`n     Processed shared mailbox count: $ProgressIndex"`n"  Checking $MailboxName mailbox"
    # Retrieve additional details using Microsoft Graph
    $User = Get-MgUser -UserId $_.ExternalDirectoryObjectId -Property AccountEnabled, DisplayName, UserPrincipalName, SignInActivity | 
            Select-Object DisplayName, UserPrincipalName, AccountEnabled, @{Name="LastSignInTime";Expression={$_.SignInActivity.LastSignInDateTime}}
    
    if ($User.AccountEnabled -eq $true) {
        $LicenseDetails = Get-MgUserLicenseDetail -UserId $User.UserPrincipalName
        $HasExchangeOnline = $false
        foreach ($License in $LicenseDetails) {
            foreach ($ServicePlan in $License.ServicePlans) {
                if ($ExchangeOnlineServicePlans -contains $ServicePlan.ServicePlanId) {
                    $HasExchangeOnline = $true
                    break
                }
            }
            if ($HasExchangeOnline) { break }
        }

        # Check for non-compliance
        if (-not $HasExchangeOnline) {
            $NonCompliantCount++

            # Gather all necessary properties
            $NonCompliantDetails = [PSCustomObject]@{
                "Shared Mailbox Name"    = $User.DisplayName
                "Primary SMTP Address"   = $_.PrimarySmtpAddress
                "Sign-In Enabled"        = "Enabled"
                "Exchange License"       = "No"
                "Last Sign-In Time"      = if ($User.LastSignInTime -eq $null) { "Never logged-in" } else { $User.LastSignInTime }
                "Creation Time"          = $_.WhenCreated
            }

            # Append to CSV
            $NonCompliantDetails | Export-Csv -Path $OutputCSV -NoTypeInformation -Append
        }
    }
}    


Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1900+ Microsoft 365 reports. ~~" -ForegroundColor Green `n

# Disconnect sessions
#Disconnect-MgGraph | Out-Null
#Disconnect-ExchangeOnline -Confirm:$false | Out-Null


# Prompt to open CSV if non-compliant mailboxes exist
if(Test-Path -Path $OutputCSV) {   
    Write-Host  " Found $NonCompliantCount non-compliant shared mailboxes." -ForegroundColor Yellow;
    Write-Host  " Output CSV file saved to: " -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCSV"  
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open the output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6) {
        Invoke-Item "$OutputCSV"
    }
}
else {
    Write-Host `n"No non-compliant shared mailboxes found." 
}