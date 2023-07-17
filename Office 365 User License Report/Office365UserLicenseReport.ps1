<#
=============================================================================================
Name:           Export Office 365 users' license report
website:        o365reports.com

~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~
1. The script uses MS Graph PowerShell and installs MS Graph PowerShell SDK (if not installed already) upon your confirmation. 
2. It can be executed with certificate-based authentication (CBA) too.
3. Exports Office 365 user license report to CSV file.
4. You can choose to either “export license report for all office 365 users” or pass an input file to get license report of specific users alone.
5. License Name is shown with its friendly name  like ‘Office 365 Enterprise E3’ rather than ‘ENTERPRISEPACK’.
6. The script can be executed with MFA enabled account too.
7. The script gives 2 output files. One with the detailed report of O365 Licensed users another with the simple details.


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
    Write-Host "Microsoft Graph Beta PowerShell module is connected successfully" -ForegroundColor Green
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
    $SKUs = Get-MgBetaUserLicenseDetail -UserId $UPN -ErrorAction SilentlyContinue
    $LicenseCount = $SKUs.count
    $count = 0
    foreach($Sku in $SKUs)  #License loop
    {
        if($FriendlyNameHash[$Sku.SkuPartNumber])
        {
            $NamePrint = $FriendlyNameHash[$Sku.SkuPartNumber]
        }
        else
        {
            $NamePrint = $Sku.SkuPartNumber
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
            $ServiceFriendlyName = $ServiceArray|Where-Object{$_.Service_Plan_Name -eq $ServiceName}
            if($ServiceFriendlyName -ne $Null)
            {
                $ServiceFriendlyName = $ServiceFriendlyName[0].ServiceFriendlyNames
            }
            else
            {
                $ServiceFriendlyName = $ServiceName
            }
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
    Write-Host "`nNote: If you encounter module related conflicts, run the script in a fresh PowerShell window." -ForegroundColor Yellow
    #Set output file
    $ExportCSV = ".\DetailedO365UserLicenseReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
    $ExportSimpleCSV = ".\SimpleO365UserLicenseReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
    #FriendlyName list for license plan and service
    try{
        $FriendlyNameHash = Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction SilentlyContinue -ErrorVariable ERR | ConvertFrom-StringData
        if($ERR -ne $null)
        {
            Write-Host $ERR -ForegroundColor Red
            CloseConnection
        }
        $ServiceArray =  Import-Csv -Path .\ServiceFriendlyName.csv 
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
            Get-MgBetaUser -UserId $item.UserPrincipalName -ErrorAction SilentlyContinue |Where-Object{$_.AssignedLicenses -ne $null} | ForEach-Object{
                Get_UsersLicenseInfo
                $LicensedUserCount++
            }
        }
    }
    #Get all licensed users
    else
    {
        Get-MgBetaUser -All  |Where-Object{$_.AssignedLicenses -ne $null} | ForEach-Object{
            Get_UsersLicenseInfo
            $LicensedUserCount++
        }
    }
    #Open output file after execution
    if((Test-Path -Path $ExportCSV) -eq "True") 
    {   Write-Host `n "Detailed report available in:" -NoNewline -ForegroundColor Yellow; Write-Host "$ExportCSV" 
        Write-Host `n "Simple report available in:" -NoNewline -ForegroundColor Yellow; Write-Host "$ExportSimpleCSV" `n 
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
        Write-Host "No data found" 
    }
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    CloseConnection
}
 . main
