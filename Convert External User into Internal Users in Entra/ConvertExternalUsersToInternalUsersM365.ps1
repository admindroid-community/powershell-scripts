<#
=============================================================================================
Name:        Bulk Convert External Users to Internal Users using PowerShell
Description: Converts external users to internal users in bulk using PowerShell while preserving access and generating detailed audit logs.
Version:     1.0
Website:     o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Supports bulk conversion of guest users to internal users using CSV input. 
2. Automatically generates UPN and password if not provided. 
3. Allows both single-user and bulk-user conversion.
4. This script automatically checks for Microsoft Graph Beta and 7Zip4Powershell modules and installs them if missing with your confirmation. 
5. Generates detailed execution logs with conversion status and exports them as a password-protected ZIP file for secure auditing.
6. The script can be executed with an MFA enabled account too. 
7. The script is compatible with certificate-based authentication (CBA).
 
For detailed script execution: https://o365reports.com/convert-external-users-into-internal-users-in-microsoft-entra/

=============================================================================================
#>

param (
    [Parameter(Mandatory, ParameterSetName = 'Bulk')]
    [string]$InputCSVFilePath,

    [Parameter(ParameterSetName = 'Single')]
    [switch]$AutoGenerateNewUPN,
    
    [Parameter(ParameterSetName = 'Single')]
    [switch]$AutoGeneratePassword,
    
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

function Connect_ToMgGraph {
    $MsGraphBetaModule = Get-Module Microsoft.Graph.Beta -ListAvailable
    if ($MsGraphBetaModule -eq $null) {
        Write-Host "`nImportant: Microsoft Graph Beta module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host "`nAre you sure you want to install Microsoft Graph Beta module? [Y] Yes [N] No"
        if ($confirm -match "[yY]") {
            Write-Host "`nInstalling Microsoft Graph Beta module..."
            Install-Module Microsoft.Graph.Beta -Scope CurrentUser -AllowClobber
            Write-Host "`nMicrosoft Graph Beta module is installed in the machine successfully" -ForegroundColor Magenta
        } else {
            Write-Host "`nExiting. `nNote: Microsoft Graph Beta module must be available in your system to run the script. Please install the module using 'Install-Module Microsoft.Graph.Beta' cmdlet." -ForegroundColor Red
            Exit
        }
    } 

    $7ZipModule = Get-Module -Name 7Zip4Powershell -ListAvailable
    if ($7ZipModule -eq $null) {
        Write-Host "`nImportant: 7Zip4Powershell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host "`nAre you sure you want to install 7Zip4Powershell module? [Y] Yes [N] No"
        if ($confirm -match "[yY]") {
            Write-Host "`nInstalling 7Zip4Powershell module..."
            Install-Module 7Zip4Powershell -Scope CurrentUser -AllowClobber -Force
            Write-Host "`n7Zip4Powershell module is installed in the machine successfully" -ForegroundColor Magenta
        } else {
            Write-Host "`nExiting. `nNote: 7Zip4Powershell module must be available in your system to run the script" -ForegroundColor Red
            Exit
        }
    } 
    Import-Module 7Zip4Powershell

    Write-Host "`nConnecting to Microsoft Graph..."
    if (($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne "")) {
        Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    } else {
        Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome
    }

    if ((Get-MgContext) -ne $null) {
        Write-Host "`nConnected to Microsoft Graph PowerShell Module."
    } else {
        Write-Host "`nFailed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
}

function Log-ScriptExecution {
    param (
        [string]$OldUserPrincipalName,
        [string]$NewUserPrincipalName,
        [string]$NewPassword,
        [Nullable[boolean]]$IsPasswordChangeRequired,
        [boolean]$Status,
        [string]$ResultDetails
    )
    $Timestamp = (Get-Date).ToLocalTime()

    $LogObject = [PSCustomObject]@{
        "Event Time"            = $Timestamp
        "Old UserPrincipalName" = $OldUserPrincipalName
        "New UserPrincipalName" = if ([string]::IsNullOrEmpty($NewUserPrincipalName)) { "-" } else { $NewUserPrincipalName }
        "New Password"          = if ([string]::IsNullOrEmpty($NewPassword)) { "-" } else { $NewPassword }
        "Is Password Change Required" = if ($null -eq $IsPasswordChangeRequired) { "-" } elseif ($true -eq $IsPasswordChangeRequired) { "Yes" } else { "No" }
        "Operation Status"      = if ($Status) { "Success" } else { "Failed" }
        "Result Details"        = if ([string]::IsNullOrEmpty($ResultDetails)) { "-" } else { $ResultDetails }
    }

    $LogObject | Export-Csv -Path $LogFilePath -NoTypeInformation -Append
}

function ValidateAndImportCsv {
    param (
        [string]$FilePath
    )

    if (-not (Test-Path $FilePath)) {
        Write-Host "Input CSV file not found: $FilePath" -ForegroundColor Red
        Exit 1
    }

    $data = Import-Csv $FilePath
    if ($data.Count -eq 0) {
        Write-Host "Input CSV file is empty." -ForegroundColor Red
        Exit 1
    }

    if ("UserId" -notin $data[0].PSObject.Properties.Name) {
        Write-Host "Missing required CSV column: UserId" -ForegroundColor Red
        Exit 1
    }

    return $data
}


function ConvertToInternalUser {
    param (
        [string]$UserId,
        [string]$NewUserPrincipalName,
        [String]$NewPassword,
        [String]$ForceChangePasswordNextSignIn
    )

    $ConversionResult = @{
        NewPassword = $null;
        ResultDetails = $null;
        NewUserPrincipalName = $null;
        OldUserPrincipalName = $null;
        IsPasswordChangeRequired = $null;
    }

    if ([string]::IsNullOrEmpty($UserId) ) {
        $ConversionResult.ResultDetails = "UserId / UserPrincipalName / Email cannot be null or empty!"
    } else {
        $ConversionResult.OldUserPrincipalName = $UserId
        try {
            $User = Get-MgBetaUser -UserId $UserId -ErrorAction Stop
        }
        catch {
            $User = Get-MgBetaUser -Filter "mail eq '$UserId'" -ConsistencyLevel eventual -Property Id,UserPrincipalName,Mail -ErrorAction Stop            
            if(-not $User){
                $ConversionResult.ResultDetails = $_.Exception.Message
                return $ConversionResult
            }
        }
        $ConversionResult.OldUserPrincipalName = $User.UserPrincipalName

        if($AutoGenerateNewUPN.IsPresent -or (-not [string]::IsNullOrEmpty($InputCSVFilePath) -and [string]::IsNullOrEmpty($NewUserPrincipalName))) {
            $LocalPart = $User.Mail.Split('@')[0]
            $DomainPart = $User.UserPrincipalName.Split('@')[1]            
            $NewUserPrincipalName = "$LocalPart@$DomainPart"
            $ConversionResult.ResultDetails = "New UserPrincipalName is auto-generated."
        }
        elseif([string]::IsNullOrEmpty($NewUserPrincipalName)) {
            $ConversionResult.ResultDetails = "New UserPrincipalName cannot be null or empty. Use -AutoGenerateNewUPN to generate one."
            return $ConversionResult
        }

        $PasswordProfile = @{
            Password = $NewPassword; ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn
        }

        if($AutoGeneratePassword.IsPresent -or (-not [string]::IsNullOrEmpty($InputCSVFilePath) -and [string]::IsNullOrEmpty($PasswordProfile.Password))) {             
            $PasswordProfile.Password = -join ((33..126) | Get-Random -Count 12 | ForEach-Object {[char]$_})
            if($ConversionResult.ResultDetails -eq "New UserPrincipalName is auto-generated.") {
                $ConversionResult.ResultDetails = "New UserPrincipalName and Password are auto-generated."
            } else {
                $ConversionResult.ResultDetails = "Password is auto-generated."
            }
        }
        elseif([string]::IsNullOrEmpty($PasswordProfile.Password)) {
            $ConversionResult.ResultDetails = "Password cannot be null or empty. Use -AutoGeneratePassword to generate one."
            return $ConversionResult
        }

        if($PasswordProfile.ForceChangePasswordNextSignIn -match "[nN]") { 
			$PasswordProfile.ForceChangePasswordNextSignIn = $false
		} else { 
			$PasswordProfile.ForceChangePasswordNextSignIn = $true
		}

        try {
            Convert-MgBetaUserExternalToInternalMemberUser -UserId $User.Id -UserPrincipalName $NewUserPrincipalName -PasswordProfile $PasswordProfile -ErrorAction Stop | Out-Null
			$ConversionResult.NewPassword = $PasswordProfile.Password
            $ConversionResult.NewUserPrincipalName = $NewUserPrincipalName
            $ConversionResult.IsPasswordChangeRequired = $PasswordProfile.ForceChangePasswordNextSignIn
        } 
        catch {
            $ConversionResult.ResultDetails = $_.Exception.Message
        }
    }
    return $ConversionResult
}

Connect_ToMgGraph
$UsersToConvert = @()
$LogFilePath = "$(Get-Location)\M365_GuestToMember_Conversion_Log_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString())).csv"
$ZipFilePath = "$(Get-Location)\M365InternalUserTypeConversionLog$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).zip"

if ([string]::IsNullOrEmpty($InputCSVFilePath)) {

    $UserId = Read-Host "`nEnter UserId/UserPrincipalName/Email of the user"
    if($AutoGenerateNewUPN.IsPresent) {
        $NewUserPrincipalName = $null  
    } else {
        $NewUserPrincipalName = Read-Host "Enter a new UserPrincipalName for the user"
    }

    if ($AutoGeneratePassword.IsPresent) {
        $NewPassword = $null
        $ForceChangePasswordNextSignIn = $true
    } else {
        $NewSecurePassword = Read-Host "Enter a strong password for the user" -AsSecureString
        $NewPassword = [System.Net.NetworkCredential]::new("", $NewSecurePassword).Password
        $ForceChangePasswordNextSignIn = Read-Host "Do you want the user to change their password on next sign-in? [Y] Yes [N] No "
    }

    $UsersToConvert = [PSCustomObject]@{
        UserId = $UserId
        NewUserPrincipalName = $NewUserPrincipalName
        NewPassword = $NewPassword
        ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn
    }
}
else {
    $UsersToConvert = ValidateAndImportCsv -FilePath $InputCSVFilePath
}

$ZipFilePassword = Read-Host "`nEnter password to protect the output log file" -AsSecureString

$SuccessCount = 0
$Total = @($UsersToConvert).Count
$Count = 0
foreach ($User in $UsersToConvert) {
    
    Write-Progress -Activity "User conversion progress" -Status "Processing user: $($User.UserId) ($SuccessCount of $Total completed)" -PercentComplete (($Count/$Total) * 100)
    $ConversionResult = ConvertToInternalUser -UserId $User.UserId -NewUserPrincipalName $User.NewUserPrincipalName -NewPassword $User.NewPassword -ForceChangePasswordNextSignIn $User.ForceChangePasswordNextSignIn

    if ($ConversionResult.NewPassword -ne $null) {
        Log-ScriptExecution -Status $true -OldUserPrincipalName $ConversionResult.OldUserPrincipalName -NewUserPrincipalName $ConversionResult.NewUserPrincipalName -NewPassword $ConversionResult.NewPassword -IsPasswordChangeRequired $ConversionResult.IsPasswordChangeRequired -ResultDetails $ConversionResult.ResultDetails
        $SuccessCount++
    }
    else {
        Log-ScriptExecution -Status $false -OldUserPrincipalName $ConversionResult.OldUserPrincipalName -ResultDetails $ConversionResult.ResultDetails
    }
    $Count++
}

Write-Progress -Activity "User conversion progress" -Status "Conversion completed: $SuccessCount of $Total users processed successfully" -PercentComplete 100

Disconnect-MgGraph | Out-Null

if(Test-Path -Path $LogFilePath) 
{
    Compress-7Zip -Path $LogFilePath -ArchiveFileName $ZipFilePath -SecurePassword $ZipFilePassword -Format Zip -EncryptionMethod ZipCrypto
    Remove-Item $LogFilePath -Force
    if (Test-Path -Path $ZipFilePath) {
        Write-Host "`nThe output log file is available in: " -NoNewline -ForegroundColor Yellow
        Write-Host "$ZipFilePath"
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open the output log file?", 0,"Open output log file", 36)
        if ($UserInput -eq 6)
        {
            Invoke-Item $ZipFilePath
        }
    }

    Write-Host `n~~ Script prepared by Admindroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green `n
}