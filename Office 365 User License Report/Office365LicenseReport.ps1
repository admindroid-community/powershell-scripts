<#
=============================================================================================
Name:           Export Office 365 users' license report
website:        o365reports.com
For detailed Script execution: https://o365reports.com/2018/12/14/export-office-365-user-license-report-powershell/
============================================================================================
#>

Param
(
 [Parameter(Mandatory = $false)]
    [string]$UserNamesFile,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)
Function ConnectMgGraphModule 
{
    $MsGraphModule =  Get-Module Microsoft.Graph -ListAvailable
    if($MsGraphModule -eq $null)
    { 
        Write-host "Important: Microsoft Graph PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install Microsoft graph module? [Y] Yes [N] No  
        if($confirm -match "[yY]") 
        { 
            Write-host "Installing Microsoft Graph PowerShell module..."
            Install-Module Microsoft.Graph -Scope CurrentUser
            Write-host "Microsoft Graph PowerShell module is installed in the machine successfully" -ForegroundColor Magenta 
        } 
        else
        { 
            Write-host "Exiting. `nNote: Microsoft Graph PowerShell module must be available in your system to run the script" -ForegroundColor Red
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
    }
    else
    {
        Connect-MgGraph -Scopes "Directory.Read.All"  -ErrorAction SilentlyContinue -Errorvariable ConnectionError |Out-Null
        if($ConnectionError -ne $null)
        {
            Write-Host $ConnectionError -Foregroundcolor Red
            Exit
        }
    }
    Write-Host "Microsoft Graph PowerShell module is connected successfully" -ForegroundColor Green
}

Function Get_UsersLicenseInfo
{
    $LicensePlanWithEnabledService=""
    $FriendlyNameOfLicensePlanWithService=""
    $UPN = $_.UserPrincipalName
    $Country = $_.Country
    if($Country -eq "")
    {
        $Country="-"
    }
    Write-Progress -Activity "`n     Exported user count:$LicensedUserCount "`n"Currently Processing:$upn"
    $SKUs = Get-MgUserLicenseDetail -UserId $UPN
    $LicenseCount = $SKUs.count
    $count = 0
    foreach($Sku in $SKUs)  #License loop
    {
        $EasyName = $FriendlyNameHash[$Sku.SkuPartNumber]
        if(!($EasyName))
        {
            $NamePrint = $Sku.SkuPartNumber
        }
        else
        {
            $NamePrint = $EasyName
        }
        #Get all services for current SKUId
        $Services = $Sku.ServicePlans
        if(($Count -gt 0) -and ($count -lt $LicenseCount))
        {
            $LicensePlanWithEnabledService = $LicensePlanWithEnabledService+","
            $FriendlyNameOfLicensePlanWithService = $FriendlyNameOfLicensePlanWithService+","
        }
        $DisabledServiceCount = 0
        $EnabledServiceCount = 0
        $serviceExceptDisabled = ""
        $FriendlyNameOfServiceExceptDisabled = ""
        foreach($Service in $Services) #Service loop
        {
            $flag = 0
            $ServiceName = $Service.ServicePlanName
            if($service.ProvisioningStatus -eq "Disabled")
            {
                $DisabledServiceCount++
            }
            else
            {
                $EnabledServiceCount++
                if($EnabledServiceCount -ne 1)
                {
                    $serviceExceptDisabled = $serviceExceptDisabled+","
                }
                $serviceExceptDisabled = $serviceExceptDisabled+$ServiceName
                $flag = 1
            }
            #Convert ServiceName to friendly name
            $ServiceFriendlyName = $ServiceArray|?{$_.Service_Plan_Name -eq $ServiceName}
            $ServiceFriendlyName = $ServiceFriendlyName[0].ServiceFriendlyNames
            if($flag -eq 1)
            {
                if($EnabledServiceCount -ne 1)
                {
                    $FriendlyNameOfServiceExceptDisabled = $FriendlyNameOfServiceExceptDisabled+","
                }
                $FriendlyNameOfServiceExceptDisabled = $FriendlyNameOfServiceExceptDisabled+$ServiceFriendlyName
            }
            $Result = [PSCustomObject]@{'DisplayName'=$_.Displayname;'UserPrinciPalName'=$UPN;'LicensePlan'=$Sku.SkuPartNumber;'FriendlyNameofLicensePlan'=$nameprint;'ServiceName'=$ServiceName;'FriendlyNameofServiceName'=$serviceFriendlyName;'ProvisioningStatus'=$service.ProvisioningStatus}
            $Result  | Export-Csv -Path $ExportCSV -Notype -Append
        }
        if($Disabledservicecount -eq 0)
        {
            $serviceExceptDisabled = "All services"
            $FriendlyNameOfServiceExceptDisabled = "All services"
        }
        $LicensePlanWithEnabledService = $LicensePlanWithEnabledService + $Sku.SkuPartNumber +"[" +$serviceExceptDisabled +"]"
        $FriendlyNameOfLicensePlanWithService = $FriendlyNameOfLicensePlanWithService+ $NamePrint + "[" + $FriendlyNameOfServiceExceptDisabled +"]"
        #Increment SKUid count
        $count++
     }
     $Output=[PSCustomObject]@{'Displayname'=$_.Displayname;'UserPrincipalName'=$UPN;Country=$Country;'LicensePlanWithEnabledService'=$LicensePlanWithEnabledService;'FriendlyNameOfLicensePlanAndEnabledService'=$FriendlyNameOfLicensePlanWithService}
     $Output | Export-Csv -path $ExportSimpleCSV -NoTypeInformation -Append
}
function CloseConnection
{
    Disconnect-MgGraph |Out-Null
    Exit
}
Function main()
{
    ConnectMgGraphModule
    Select-MgProfile beta
    #Set output file
    $ExportCSV = ".\DetailedO365UserLicenseReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
    $ExportSimpleCSV = ".\SimpleO365UserLicenseReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
    #FriendlyName list for license plan and service
    try{
        $FriendlyNameHash = Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction SilentlyContinue | ConvertFrom-StringData
        $ServiceArray =  Import-Csv -Path .\ServiceFriendlyName.csv -ErrorAction SilentlyContinue 
    }
    catch
    {
        Write-Host $_.Exception.message -ForegroundColor Red
        CloseConnection
    }
    #Get licensed user
    $LicensedUserCount = 0
    #Check for input file/Get users from input file
    if($UserNamesFile -ne "")
    {
        #We have an input file, read it into memory
        $UserNames = @()
        $UserNames = Import-Csv -Header "UserPrincipalName" $UserNamesFile
        foreach($item in $UserNames)
        {
            Get-MgUser -UserId $item.UserPrincipalName |where{$_.AssignedLicenses -ne $null} | Foreach{
            Get_UsersLicenseInfo
            $LicensedUserCount++
            }
        }
    }
    #Get all licensed users
    else
    {
        Get-MgUser -All | where{$_.AssignedLicenses -ne $null} | Foreach{
        Get_UsersLicenseInfo
        $LicensedUserCount++}
    }
    #Open output file after execution
    if((Test-Path -Path $ExportCSV) -eq "True") 
    {
        Write-Host Detailed report available in: $ExportCSV
        Write-host Simple report available in: $ExportSimpleCSV
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open output files?",` 0,"Open Files",4)
        if($UserInput -eq 6)
        {
            Invoke-Item $ExportCSV
            Invoke-Item $ExportSimpleCSV
        }
    }
    else
    {
        Write-Host "No data found" -ForegroundColor Red
    }
    CloseConnection
}
 . main
