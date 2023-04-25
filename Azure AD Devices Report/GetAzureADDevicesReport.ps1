<#
=============================================================================================
Name:           Get Azure AD Devices Report Using PowerShell
Description:    This script gives detailed information on all Azure AD devices
Version:        1.0
Website:        o365reports.com
For detailed script execution: https://o365reports.com/2023/04/18/get-azure-ad-devices-report-using-powershell/
============================================================================================
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
$MsGraphModule =  Get-Module Microsoft.Graph -ListAvailable
if($MsGraphModule -eq $null)
{ 
    Write-host "Important: Microsoft Graph Powershell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
    $confirm = Read-Host Are you sure you want to install Microsoft Graph Powershell module? [Y] Yes [N] No  
    if($confirm -match "[yY]") 
    { 
        Write-host "Installing Microsoft Graph Powershell module..."
        Install-Module Microsoft.Graph -Scope CurrentUser
        Write-host "Microsoft Graph Powershell module is installed in the machine successfully" -ForegroundColor Magenta 
    } 
    else
    { 
        Write-host "Exiting. `nNote: Microsoft Graph Powershell module must be available in your system to run the script" -ForegroundColor Red
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
Write-Host "Microsoft Graph Powershell module is connected successfully" -ForegroundColor Green
Select-MgProfile beta
function CloseConnection
{
    Disconnect-MgGraph |  Out-Null
    Exit
}
$OutputCsv =".\AzureDeviceReport_$((Get-Date -format MMM-dd` hh-mm-ss` tt).ToString()).csv" 
$Report=""
$FilterCondition = @()
$DeviceInfo = Get-MgDevice -All
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
            $BitLockerKeys = Get-MgInformationProtectionBitlockerRecoveryKey -Filter "DeviceId eq '$($Device.DeviceId)'" -ErrorAction SilentlyContinue -ErrorVariable Err
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
    $DeviceOwners = Get-MgDeviceRegisteredOwner -DeviceId $Device.Id -All |Select-Object -ExpandProperty AdditionalProperties
    $DeviceUsers = Get-MgDeviceRegisteredUser -DeviceId $Device.Id -All |Select-Object -ExpandProperty AdditionalProperties
    $DeviceMemberOf = Get-MgDeviceMemberOf -DeviceId $Device.Id -All |Select-Object -ExpandProperty AdditionalProperties
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
    Write-Host "The Output file availble in $outputCsv" -ForegroundColor Green 
    $prompt = New-Object -ComObject wscript.shell    
    $UserInput = $prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)    
    if ($UserInput -eq 6)    
    {    
        Invoke-Item "$OutputCsv"  
        Write-Host "Report generated successfully"  -ForegroundColor Green 
    }
} 
else
{
    Write-Host "No devices found" -ForegroundColor Red
}
CloseConnection