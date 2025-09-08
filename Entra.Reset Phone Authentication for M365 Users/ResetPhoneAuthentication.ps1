<#
=========================================================================================
Name:           Delete phone authentication for Microsoft 365 users
Version:        1.0
Website:        blog.admindroid.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script utilizes Microsoft Graph PowerShell and installs it (if not already installed) upon your confirmation.
2. The script covers 7+ usecases to delete Phone MFA for Microsoft 365 users more granularly.
3. Users scope
       - Single user
       - Bulk users (input CSV)
       - All users
       - Admin accounts
       - Guest accounts 
       - Licensed users
       - Disabled users
4. This scripts supports deletion of all types of Phone MFA (primary mobile, alternate mobile and Office phone).     
5. Exports log file.
6. It can be executed with a MFA-enabled account too.     
7. Supports certificate-based authentication. 
8. The script is scheduler-friendly.

For detailed script execution: https://blog.admindroid.com/delete-phone-authentication-for-microsoft-365-users/ 

=========================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [ValidateSet(
        'Primary mobile',
        'Alternate mobile',
        'Office'
    )]
    [string]$ResetPhoneMFA,
    [string]$UserId,
    [string]$CsvFilePath,
    [switch]$AllUsers,
    [switch]$AdminsOnly,
    [switch]$GuestUsersOnly,
    [switch]$LicensedUsersOnly,
    [switch]$DisabledUsersOnly,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint

)

# Function to install and connect to Microsoft Graph
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
        Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.ReadWrite.All" -NoWelcome 
    }

    # Verify connection
    if ((Get-MgContext) -ne $null) {
        Write-Host "Connected to Microsoft Graph PowerShell using account: $((Get-MgContext).Account)`n" -ForegroundColor Yellow
    } else {
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
}

# Function to log details of authentication method removal
function Log-MFAReset {
    param (
        [string]$UserId,
        [string]$AuthMethodType,
        [boolean]$Status
    )
    $Timestamp = (Get-Date).ToLocalTime()
    if ($Status) {  
        $LogEntry = "$UserId's $AuthMethodType MFA method has been reset successfully."
    } else {
        $LogEntry = "Error occurred during $AuthMethodType MFA method reset for $UserId. Please check the default authentication method settings."
    }
    Add-Content -Path $LogFilePath -Value $LogEntry
}

# Function to reset MFA authentication method for users
function Reset-MFA {
    param(
        [string]$UserId,
        [object[]]$UserAuthenticationDetail
    )
    
    $MethodType = $UserAuthenticationDetail.PhoneType
    $FriendlyAuthName = $AuthMethods.Keys | Where-Object { $AuthMethods[$_] -eq $MethodType }

    # Perform removal based on the method type
    $Script:ResetStatus = Remove-MgUserAuthenticationPhoneMethod -UserId $UserId -PhoneAuthenticationMethodId $UserAuthenticationDetail.Id -PassThru

    # Log the MFA reset based on the MFA reset status
    Log-MFAReset -UserId $UserId -AuthMethodType $FriendlyAuthName -Status $Script:ResetStatus
}

# Function to call reset MFA for each users in array
function Reset-MfaForUsers {
    param(
        [string[]]$Users,
        [string]$SpecificAuthMethod
    )
    
    $AuthMethods = @{
        "Primary mobile" = "mobile";
        "Alternate mobile" = "alternateMobile";
        "Office" = "office";
    }
    foreach ($User in $Users) {
        if (-not [string]::IsNullOrEmpty($SpecificAuthMethod)) {
            $UserAuthenticationDetails = Get-MgUserAuthenticationPhoneMethod -UserId $User | Where-Object { $_.PhoneType -eq $AuthMethods[$SpecificAuthMethod] } 
        } else {
            $UserAuthenticationDetails = Get-MgUserAuthenticationPhoneMethod -UserId $User 
        }
        foreach ($UserAuthenticationDetail in $UserAuthenticationDetails) {
            $Script:ResetStatus = $false
            Reset-MFA -UserId $User -UserAuthenticationDetail $UserAuthenticationDetail
        }
    }
}

# Connecting to the Microsoft Graph PowerShell Module
Connect_ToMgGraph

# Define log file path and get users
$LogFilePath = "$(Get-Location)\Phone_MFA_Reset_Log_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).txt"
$Users = Get-MgUser -All -Property AccountEnabled, AssignedLicenses, UserType, UserPrincipalName | Where-Object { $_.UserPrincipalName -ne ((Get-MgContext).Account) } | Select-Object -Property AccountEnabled, AssignedLicenses, UserType, UserPrincipalName


if (-not [string]::IsNullOrEmpty($CsvFilePath)) {
    # Load users from the CSV file
    $Users = Import-CSV -Path $CsvFilePath 
    $Users.Name | ForEach-Object { Reset-MfaForUsers -Users $_ -SpecificAuthMethod $ResetPhoneMFA }
}
elseif (-not [string]::IsNullOrEmpty($UserId)) {
    Reset-MfaForUsers -Users $UserId -SpecificAuthMethod $ResetPhoneMFA 
}
elseif (!$DisabledUsersOnly -and !$LicensedUsersOnly -and !$GuestUsersOnly -and !$AdminsOnly -and !$AllUsers) {
    $UserId = Read-Host "Enter the User ID or UPN of a User to Reset MFA"
    Reset-MfaForUsers -Users $UserId -SpecificAuthMethod $ResetPhoneMFA 
}
elseif ($AllUsers.IsPresent) {
    $Users | ForEach-Object { Reset-MfaForUsers -Users $_.UserPrincipalName -SpecificAuthMethod $ResetPhoneMFA }
}
elseif ($DisabledUsersOnly.IsPresent) {
    $Users | Where-Object { $_.AccountEnabled -eq $false } | ForEach-Object { Reset-MfaForUsers -Users $_.UserPrincipalName -SpecificAuthMethod $ResetPhoneMFA } 
}
elseif ($LicensedUsersOnly.IsPresent) {
    $Users | Where-Object { $_.AssignedLicenses } | ForEach-Object { Reset-MfaForUsers -Users $_.UserPrincipalName -SpecificAuthMethod $ResetPhoneMFA }
}
elseif ($GuestUsersOnly.IsPresent) {
    $Users | Where-Object { $_.UserType -eq "Guest" } | ForEach-Object { Reset-MfaForUsers -Users $_.UserPrincipalName -SpecificAuthMethod $ResetPhoneMFA }
}
elseif ($AdminsOnly.IsPresent) {
    $Users | Where-Object { Get-MgUserTransitiveMemberOfAsDirectoryRole -UserId $_.UserPrincipalName } | ForEach-Object { Reset-MfaForUsers -Users $_.UserPrincipalName -SpecificAuthMethod $ResetPhoneMFA }
}


Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1900+ Microsoft 365 reports. ~~" -ForegroundColor Green `n


# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

if((Test-Path -Path $LogFilePath) -eq "True") {   
    Write-Host  " The MFA reset log file available in: " -NoNewline -ForegroundColor Yellow; Write-Host "$LogFilePath"  
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open the log file?",` 0,"Open Log File",4)
    if ($UserInput -eq 6) {
        Invoke-Item "$LogFilePath"
    }
}
else {
    Write-Host "`nNo user(s) found with specific criteria." -ForegroundColor Yellow
}
