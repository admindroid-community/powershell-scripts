<#
=============================================================================================
Name:           Export Azure Device Report using MS Graph PowerShell
Description:    This script exports Microsoft 365 Azure AD devices to CSV
Version:        2.0
website:        o365reports.com


Script Highlights 
1.The script can be executed with MFA-enabled accounts too. 
2.Exports output to CSV. 
3.Automatically installs the Microsoft Graph PowerShell module in your PowerShell environment upon your confirmation. 
4.Supports the method of certificate-based authentication. 
5.The script lists all the Azure AD devices of your organization. That too customization of reports is possible according to the major device types like managed, enabled, disabled etc. 

For detailed script execution: https://o365reports.com/2023/04/18/get-azure-ad-devices-report-using-powershell/
#>




## If you execute via CBA, then your application required "Directory.Read.All" application permissions.
Param
(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$EnabledDevice,
    [switch]$DisabledDevice,
    [Int]$InactiveDays,
    [switch]$ManagedDevice,
    [switch]$DevicesWithBitLockerKey
)
$MsGraphBetaModule =  Get-Module Microsoft.Graph.Beta -ListAvailable
if($MsGraphBetaModule -eq $null)
{ 
    Write-host "Important: Microsoft Graph Beta module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
    $confirm = Read-Host Are you sure you want to install Microsoft Graph Beta module? [Y] Yes [N] No  
    if($confirm -match "[yY]") 
    { 
        Write-host "Installing Microsoft Graph Beta module..."
        Install-Module Microsoft.Graph.Beta -Scope CurrentUser -AllowClobber
        Write-host "Microsoft Graph Beta module is installed in the machine successfully" -ForegroundColor Magenta 
    } 
    else
    { 
        Write-host "Exiting. `nNote: Microsoft Graph Beta module must be available in your system to run the script" -ForegroundColor Red
        Exit 
    } 
}
if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
{  
    Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction SilentlyContinue -ErrorVariable ConnectionError|Out-Null
    if($ConnectionError -ne $null)
    {    
        Write-Host $ConnectionError -Foregroundcolor Red
        Exit
    }
    $Certificate = (Get-MgContext).CertificateThumbprint
    Write-Host "Note: You don't get device with bitlocker key info while using certificate based authentication. If you want to get bitlocker key enabled devices, then you can connect graph using credentials(User interaction based authentication)" -ForegroundColor Yellow
}
else
{
    Connect-MgGraph -Scopes "Directory.Read.All,BitLockerKey.Read.All"  -ErrorAction SilentlyContinue -Errorvariable ConnectionError |Out-Null
    if($ConnectionError -ne $null)
    {
        Write-Host "$ConnectionError" -Foregroundcolor Red
        Exit
    }
}
Write-Host "Microsoft Graph Beta Powershell module is connected successfully" -ForegroundColor Green
Write-Host "`nNote: If you encounter module related conflicts, run the script in a fresh Powershell window."
function CloseConnection
{
    Disconnect-MgGraph |  Out-Null
    Exit
}
$OutputCsv =".\AzureDeviceReport_$((Get-Date -format MMM-dd` hh-mm-ss` tt).ToString()).csv" 
$Report=""
$FilterCondition = @()
$DeviceInfo = Get-MgBetaDevice -All
if($DeviceInfo -eq $null)
{
    Write-Host "You have no devices enrolled in your Azure AD" -ForegroundColor Red
    CloseConnection
}
if($EnabledDevice.IsPresent)
{
    $DeviceInfo = $DeviceInfo | Where-Object {$_.AccountEnabled -eq $True}
}
elseif($DisabledDevice.IsPresent)
{
    $DeviceInfo = $DeviceInfo | Where-Object {$_.AccountEnabled -eq $False}
}
if($ManagedDevice.IsPresent)
{
    $DeviceInfo = $DeviceInfo | Where-Object {$_.IsManaged -eq $True}
}
$TimeZone = (Get-TimeZone).Id
Foreach($Device in $DeviceInfo){
    Write-Progress -Activity "Fetching devices: $($Device.DisplayName)"
    $LastSigninActivity = "-"
    if(($Device.ApproximateLastSignInDateTime -ne $null))
    {
        $LastSigninActivity = (New-TimeSpan -Start $Device.ApproximateLastSignInDateTime).Days
    }
    if($Certificate -eq $null)
    {
        $BitLockerKeyIsPresent = "No"
        try {
            $BitLockerKeys = Get-MgBetaInformationProtectionBitlockerRecoveryKey -Filter "DeviceId eq '$($Device.DeviceId)'" -ErrorAction SilentlyContinue -ErrorVariable Err
            if($Err -ne $null)
            {
                Write-Host $Err -ForegroundColor Red
                CloseConnection
            }
        }
        catch
        {
            Write-Host $_.Exception.Message -ForegroundColor Red
            CloseConnection
        }
        if($BitLockerKeys -ne $null)
        {
            $BitLockerKeyIsPresent = "Yes"
        }
        if($DevicesWithBitLockerKey.IsPresent)
        {
            if($BitLockerKeyIsPresent -eq "No")
            {
                Continue
            }
        }
    }
    if($InactiveDays -ne "")
    {
        if(($Device.ApproximateLastSignInDateTime -eq $null))
        {
            Continue
        }
        if($LastSigninActivity -le $InactiveDays) 
        {
            continue
        }
    }
    $DeviceOwners = Get-MgBetaDeviceRegisteredOwner -DeviceId $Device.Id -All |Select-Object -ExpandProperty AdditionalProperties
    $DeviceUsers = Get-MgBetaDeviceRegisteredUser -DeviceId $Device.Id -All |Select-Object -ExpandProperty AdditionalProperties
    $DeviceMemberOf = Get-MgBetaDeviceMemberOf -DeviceId $Device.Id -All |Select-Object -ExpandProperty AdditionalProperties
    $Groups = $DeviceMemberOf|Where-Object {$_.'@odata.type' -eq '#microsoft.graph.group'}
    $AdministrativeUnits = $DeviceMemberOf|Where-Object{$_.'@odata.type' -eq '#microsoft.graph.administrativeUnit'}
    if($Device.TrustType -eq "Workplace")
    {
        $JoinType = "Azure AD registered"
    }
    elseif($Device.TrustType -eq "AzureAd")
    {
        $JoinType = "Azure AD joined"
    }
    elseif($Device.TrustType -eq "ServerAd")
    {
        $JoinType = "Hybrid Azure AD joined"
    }
    
    if($Device.ApproximateLastSignInDateTime -ne $null)
    {
        $LastSigninDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($Device.ApproximateLastSignInDateTime,$TimeZone) 
        $RegistrationDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($Device.RegistrationDateTime,$TimeZone)
    }
    else
    {
        $LastSigninDateTime = "-"
        $RegistrationDateTime = "-"
    }
    $ExtensionAttributes = $Device.ExtensionAttributes
    $AttributeArray = @()
    $Attributes = $ExtensionAttributes.psobject.properties |Where-Object {$_.Value -ne $null -and $_.Name -ne "AdditionalProperties"}| select Name,Value
    Foreach($Attribute in $Attributes)
    {
        $AttributeArray+=$Attribute.Name+":"+$Attribute.Value
    }
    $ExportResult = @{'Name'                 =$Device.DisplayName
                    'Enabled'                ="$($Device.AccountEnabled)"
                    'Operating System'       =$Device.OperatingSystem
                    'OS Version'             =$Device.OperatingSystemVersion
                    'Join Type'              =$JoinType
                    'Owners'                 =(@($DeviceOwners.userPrincipalName) -join ',')
                    'Users'                  =(@($DeviceUsers.userPrincipalName)-join ',')
                    'Is Managed'             ="$($Device.IsManaged)"
                    'Management Type'        =$Device.ManagementType
                    'Is Compliant'           ="$($Device.IsCompliant)"
                    'Registration Date Time' =$RegistrationDateTime
                    'Last SignIn Date Time'  =$LastSigninDateTime
                    'InActive Days'           =$LastSigninActivity
                    'Groups'                 =(@($Groups.displayName) -join ',')
                    'Administrative Units'   =(@($AdministrativeUnits.displayName) -join ',')
                    'Device Id'              =$Device.DeviceId
                    'Object Id'              =$Device.Id
                    'BitLocker Encrypted'    =$BitLockerKeyIsPresent
                    'Extension Attributes'   =(@($AttributeArray)| Out-String).Trim()
                    }
    $Results = $ExportResult.GetEnumerator() | Where-Object {$_.Value -eq $null -or $_.Value -eq ""} 
    Foreach($Result in $Results){
        $ExportResult[$Result.Name] = "-"
    }
    $Report = [PSCustomObject]$ExportResult
    if($Certificate -eq $null)
    {
        $Report|Select 'Name','Enabled','Operating System','OS Version','Join Type','Owners','Users','Is Managed','Management Type','Is Compliant','Registration Date Time','Last SignIn Date Time','InActive Days','Groups','Administrative Units','Device Id','Object Id','BitLocker Encrypted','Extension Attributes' | Export-csv -path $OutputCsv -NoType -Append  
    }
    else
    {
        $Report|Select 'Name','Enabled','Operating System','OS Version','Join Type','Owners','Users','Is Managed','Management Type','Is Compliant','Registration Date Time','Last SignIn Date Time','InActive Days','Groups','Administrative Units','Device Id','Object Id','Extension Attributes' | Export-csv -path $OutputCsv -NoType -Append          
    }
}
if((Test-Path -Path $OutputCsv) -eq "True") 
{ 
     Write-Host `n "The Output file availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$outputCsv" `n 
    $prompt = New-Object -ComObject wscript.shell    
    $UserInput = $prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)    
    if ($UserInput -eq 6)    
    {    
        Invoke-Item "$OutputCsv"  
        Write-Host "Report generated successfully"  
    }
} 
else
{
    Write-Host "No devices found"
}

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
CloseConnection
