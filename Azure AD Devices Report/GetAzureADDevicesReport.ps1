<#
=============================================================================================
Name:           Export Entra Device Report using MS Graph PowerShell
Description:    This script exports all Microsoft 365 devices to CSV
Version:        3.0
website:        o365reports.com

~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~
1. Exports all Azure AD devices in your organization.
2. Automatically checks the Microsoft Graph PowerShell module and installs it upon your confirmation if it’s missing.
3. Enables filtering based on following device registration types:
    -> Entra registered
    -> Entra joined
    -> Entra Hybrid joined
4. Generates a report that retrieves enabled devices alone.
5. Helps export disabled devices alone.
6. Find inactive devices based on inactive days.
7. Includes ownership filtering to show company, personal, unknown devices.
8. List all devices that have BitLocker key.
9. The script can be executed with MFA-enabled accounts too.
10. Facilitates filtering devices by users, owners, and groups.
11. Exports the report results in CSV format.
12. Allows exporting devices that match the selected filters.
    -> Device management status – [ Managed / Unmanaged ]
    -> Device compliance status – [ Compliant / Non-compliant ]
    -> Device rooted state – [ Rooted / Nonrooted ]
13. Compatible with certificate-based authentication (CBA).
14. Filter devices by profile types such as IoT, Printer, Secure VM, Shared device, and Registered device.


For detailed script execution: https://o365reports.com/2023/04/18/get-azure-ad-devices-report-using-powershell/

~~~~~~~~~~~
Change Log:
~~~~~~~~~~~
V1.0 (Apr 25, 2023) – File created.  
V2.0 (Jul 17, 2024) – Updated to use MS Graph beta PowerShell module instead of 'Select-MgProfile beta' to load beta commands.  
V3.0 (Jul 03, 2025) – Upgraded from the 'MS Graph beta module' to Microsoft Graph module and added additional filters such as device join type, profile type, compliance status, and more.

============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [Int]$InactiveDays,
    [string[]]$Users,
    [string[]]$Owners,
    [string[]]$Groups,
    [ValidateSet("Enabled", "Disabled")]
    [string]$DeviceStatus,
    [ValidateSet("Managed", "Unmanaged")]
    [string]$ManagementStatus,
    [ValidateSet("Compliant", "NonCompliant")]
    [string]$ComplianceStatus,
    [ValidateSet("Rooted", "NonRooted")]
    [string]$RootedStatus,
    [ValidateSet("RegisteredDevice", "SecureVM", "Printer", "Shared", "IoT")]
    [string[]]$ProfileType,
    [ValidateSet("Entra registered", "Entra joined", "Entra hybrid joined")]
    [string[]]$JoinType,
    [ValidateSet("Company", "Personal", "Unknown")]
    [string[]]$DeviceOwnership,
    [switch]$DevicesWithBitLockerKey
)

# Check if Microsoft Graph module is installed
$MsGraphModule =  Get-Module Microsoft.Graph -ListAvailable
if($MsGraphModule -eq $null)
{ 
    Write-host "Important: Microsoft Graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
    $confirm = Read-Host Are you sure you want to install Microsoft Graph module? [Y] Yes [N] No  
    if($confirm -match "[yY]") { 
        Write-host "Installing Microsoft Graph module..."
        Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber
        Write-host "Microsoft Graph module is installed in the machine successfully" -ForegroundColor Magenta 
    } 
    else { 
        Write-host "Exiting. `nNote: Microsoft Graph module must be available in your system to run the script" -ForegroundColor Red
        Exit 
    } 
}

Write-Host "`nConnecting to Microsoft Graph..."

if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
{  
    Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction SilentlyContinue -ErrorVariable ConnectionError | Out-Null
    if($ConnectionError -ne $null) {    
        Write-Host $ConnectionError -Foregroundcolor Red
        Exit
    }
    Write-Host "Connected to Microsoft Graph PowerShell using certificate-based authentication."
}
else
{
    Connect-MgGraph -Scopes "Directory.Read.All,BitLockerKey.Read.All"  -ErrorAction SilentlyContinue -Errorvariable ConnectionError | Out-Null
    if($ConnectionError -ne $null) {
        Write-Host "$ConnectionError" -Foregroundcolor Red
        Exit
    }
    Write-Host "Connected to Microsoft Graph PowerShell."
}

$Location = Get-Location
$CurrentDate = Get-Date
$TimeZone = (Get-TimeZone).Id
$OutputCsv = "$Location\EntraDevicesReport_$($CurrentDate.ToString('yyyy-MMM-dd-ddd hh-mm-ss tt')).csv"
$Report=""
$PrintedLogs=0

$ManagedDevices = Get-MgDeviceManagementManagedDevice | Select-Object AzureAdDeviceId, SerialNumber

Get-MgDevice -All | ForEach-Object {
    Write-Progress -Activity "Fetching devices: $($_.DisplayName)"
    $LastSigninActivity = "-"
    $TrustType = "" 

    if(($_.ApproximateLastSignInDateTime -ne $null)) {
        $LastSigninActivity = (New-TimeSpan -Start $_.ApproximateLastSignInDateTime).Days
    }

    $BitLockerKeyIsPresent = "No"
    try {
        $BitLockerKeys = Get-MgInformationProtectionBitlockerRecoveryKey -Filter "DeviceId eq '$($_.DeviceId)'" -ErrorAction SilentlyContinue -ErrorVariable Err
        if($Err -ne $null) {
            Write-Host $Err -ForegroundColor Red
            CloseConnection
        }
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        CloseConnection
    }

    if($BitLockerKeys -ne $null) { $BitLockerKeyIsPresent = "Yes" }

    if($DevicesWithBitLockerKey.IsPresent) {
        if($BitLockerKeyIsPresent -eq "No") { Continue }
    }

    if($InactiveDays -ne "") {
        if(($_.ApproximateLastSignInDateTime -eq $null)) { Continue }
        if($LastSigninActivity -le $InactiveDays) { continue }
    }

    $SerialNumber = ""
    if ($_.IsManaged) {
        $ManagedDeviceId = $_.DeviceId
        $SerialNumber = ($ManagedDevices | Where-Object { $_.AzureAdDeviceId -eq $ManagedDeviceId }).SerialNumber
    }

    $DeviceOwners = Get-MgDeviceRegisteredOwner -DeviceId $_.Id -All | Select-Object -ExpandProperty AdditionalProperties
    $DeviceUsers = Get-MgDeviceRegisteredUser -DeviceId $_.Id -All | Select-Object -ExpandProperty AdditionalProperties
    $DeviceMemberOf = Get-MgDeviceMemberOf -DeviceId $_.Id -All | Select-Object -ExpandProperty AdditionalProperties
    $DeviceGroups = $DeviceMemberOf | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group'}
    $AdministrativeUnits = $DeviceMemberOf | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.administrativeUnit'}

    if ($_.TrustType -eq "Workplace") { $TrustType = "Entra registered" }
    elseif ($_.TrustType -eq "AzureAd") { $TrustType = "Entra joined" }
    elseif ($_.TrustType -eq "ServerAd") { $TrustType = "Entra hybrid joined" }
    
    if ($_.ApproximateLastSignInDateTime -ne $null) {
        $LastSigninDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.ApproximateLastSignInDateTime,$TimeZone) 
        $RegistrationDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.RegistrationDateTime,$TimeZone)
    } 
    else {
        $LastSigninDateTime = "-"
        $RegistrationDateTime = "-"
    }

    if ($_.ComplianceExpirationDateTime -ne $null) {
        $ComplianceExpirationDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.ComplianceExpirationDateTime,$TimeZone)
    } else { 
        $ComplianceExpirationDateTime = "-" 
    }

    $ExtensionAttributes = $_.ExtensionAttributes
    $AttributeArray = @()
    $ExtensionAttributes.psobject.properties | Where-Object {$_.Value -ne $null -and $_.Name -ne "AdditionalProperties"} | select Name, Value | ForEach-Object { $AttributeArray+=$_.Name+":"+$_.Value }

    $Print = 1

    # Apply filters based on the param values...
    if ($DeviceStatus -eq "Enabled" -and $_.AccountEnabled -ne $true) { $Print = 0 }
    elseif ($DeviceStatus -eq "Disabled" -and $_.AccountEnabled -ne $false) { $Print = 0 }
    
    if ($ManagementStatus -eq "Managed" -and $_.IsManaged -ne $true) { $Print = 0 }
    elseif ($ManagementStatus -eq "Unmanaged" -and $_.IsManaged -ne $false) { $Print = 0 }
    
    if ($ComplianceStatus -eq "Compliant" -and $_.IsCompliant -ne $true) { $Print = 0 }
    elseif ($ComplianceStatus -eq "NonCompliant" -and $_.IsCompliant -ne $false) { $Print = 0 }
    
    if ($RootedStatus -eq "Rooted" -and $_.IsRooted -ne $true) { $Print = 0 }
    elseif ($RootedStatus -eq "NonRooted" -and $_.IsRooted -ne $false) { $Print = 0 }
    
    if (!([string]::IsNullOrEmpty($ProfileType)) -and ($_.ProfileType -notin $ProfileType)) { $Print = 0 }
    if (!([string]::IsNullOrEmpty($JoinType)) -and ($TrustType -notin $JoinType)) { $Print = 0 }
    if (!([string]::IsNullOrEmpty($DeviceOwnership)) -and ($_.DeviceOwnership -notin $DeviceOwnership)) { $Print = 0 }
    if (!([string]::IsNullOrEmpty($Users)) -and ($DeviceUsers.Where({ $Users -contains $_.userPrincipalName }, 'First').Count -eq 0)) { $Print = 0 }
    if (!([string]::IsNullOrEmpty($Owners)) -and ($DeviceOwners.Where({ $Owners -contains $_.userPrincipalName }, 'First').Count -eq 0)) { $Print = 0 }
    if (!([string]::IsNullOrEmpty($Groups)) -and ($DeviceGroups.Where({ $Groups -contains $_.displayName }, 'First').Count -eq 0)) { $Print = 0 }

    $ExportResult = @{'Name'                 = $_.DisplayName
                    'Enabled'                = "$($_.AccountEnabled)"
                    'Operating System'       = $_.OperatingSystem
                    'OS Version'             = $_.OperatingSystemVersion
                    'Join Type'              = $TrustType
                    'Is Managed'             = "$($_.IsManaged)"
                    'Owners'                 = (@($DeviceOwners.userPrincipalName) -join ',')
                    'Users'                  = (@($DeviceUsers.userPrincipalName)-join ',')
                    'Management Type'        = $_.ManagementType
                    'Enrollment Type'        = $_.EnrollmentType
                    'Profile Type'           = $_.ProfileType
                    'Model'                  = $_.Model
                    'Serial Number'           = $SerialNumber
                    'Device Ownership'       = "$($_.DeviceOwnership)"
                    'Is Compliant'           = "$($_.IsCompliant)"
                    'Is Rooted'              = "$($_.IsRooted)"
                    'Registration Date Time' = $RegistrationDateTime
                    'Last SignIn Date Time'  = $LastSigninDateTime
                    'Compliance Expiration Date Time' = $ComplianceExpirationDateTime
                    'InActive Days'          = $LastSigninActivity
                    'Groups'                 = (@($DeviceGroups.displayName) -join ',')
                    'Administrative Units'   = (@($AdministrativeUnits.displayName) -join ',')
                    'Object Id'              = $_.Id
                    'Device Id'              = $_.DeviceId
                    'BitLocker Encrypted'    = $BitLockerKeyIsPresent
                    'Extension Attributes'   = (@($AttributeArray) | Out-String).Trim()
                    }

    $Results = $ExportResult.GetEnumerator() | Where-Object {$_.Value -eq $null -or $_.Value -eq ""} 
    Foreach($Result in $Results) {
        $ExportResult[$Result.Name] = "-"
    }

    $Report = [PSCustomObject]$ExportResult
    if($Print -eq 1) {
       $PrintedLogs++
       $Report | Select 'Name','Enabled','Operating System','OS Version','Model','Serial Number','Join Type','Is Managed','Owners','Users','Management Type','Enrollment Type','Profile Type','Device Ownership','Is Compliant','Is Rooted','Registration Date Time','Last SignIn Date Time','InActive Days','Groups','Administrative Units','Object Id','Device Id','BitLocker Encrypted','Extension Attributes' | Export-csv -path $OutputCsv -NoType -Append          
    }
}

#Disconnect the session after execution
Disconnect-MgGraph | Out-Null

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n

#Open output file after execution
if((Test-Path -Path $OutputCsv) -eq "True") { 
    Write-Host " Exported report has $PrintedLogs device records." 
    Write-Host `n "The Output file availble in: " -NoNewline -ForegroundColor Yellow; Write-Host "$outputCsv" `n 
    $prompt = New-Object -ComObject wscript.shell    
    $UserInput = $prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)    
    if ($UserInput -eq 6) {    
        Invoke-Item "$OutputCsv"
    }
} 
else {
    Write-Host "No devices found"
}