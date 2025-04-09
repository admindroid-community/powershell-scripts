<#
=============================================================================================
Name:           Manage Microsoft 365 licenses using MS Graph PowerShell
Description:    This script can perform 10+ Office 365 reporting and management activities
website:        o365reports.com

Script Highlights :
~~~~~~~~~~~~~~~~~

1.	The script uses MS Graph PowerShell module.
2.	Generates 5 Office 365 license reports.
3.	Allows you to perform 6 license management actions that include adding or removing licenses in bulk.
4.	License Name is shown with its friendly name like ‘Office 365 Enterprise E3’ rather than ‘ENTERPRISEPACK’.
5.	Automatically installs MS Graph PowerShell module (if not installed already) upon your confirmation.
6.	The script can be executed with an MFA enabled account too.
7.	Exports the report result to CSV.
8.	Exports license assignment and removal log file.


Change Log
~~~~~~~~~~
  V1.0 (Sep 08, 2022) - File created
  V2.0 (Mar 10, 2025)  - Upgraded from MS Graph beta to production version
  V2.1 (Mar 21, 2025)  - Feature break due to module upgrade fixed.
  V2.2 (Mar 26, 2025) - Used 'Property' param to retrive user properties.
  V2.3 (Apr 05, 2025) - Updated license friednly name with latest changes and converted it as CSV file


For detailed Script execution: https://o365reports.com/2022/09/08/manage-365-licenses-using-ms-graph-powershell
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$LicenseName,
    [string]$LicenseUsageLocation,
    [int]$Action,
    [switch]$MultipleActionsMode
)

function Connect_MgGraph {
    $MsGraphBetaModule =  Get-Module Microsoft.Graph -ListAvailable
    if($MsGraphBetaModule -eq $null)
    { 
        Write-host "Important: Microsoft Graph PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install Microsoft Graph PowerShell module? [Y] Yes [N] No  
        if($confirm -match "[yY]") 
        { 
            Write-host "Installing Microsoft Graph PowerShell module..."
            Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber
            Write-host "Microsoft Graph PowerShell module is installed in the machine successfully" -ForegroundColor Magenta 
        } 
        else
        { 
            Write-host "Exiting. `nNote: Microsoft Graph PowerShell module must be available in your system to run the script" -ForegroundColor Red
            Exit 
        } 
    }
    Write-Progress "Importing Required Modules..."
    Import-Module -Name Microsoft.Graph.Identity.DirectoryManagement
    Import-Module -Name Microsoft.Graph.Users
    Import-Module -Name Microsoft.Graph.Users.Actions
    Write-Progress "Connecting MgGraph Module..."
    Connect-MgGraph -Scopes "Directory.ReadWrite.All" -NoWelcome
}
Function Open_OutputFile {
    #Open output file after execution 
    if ((Test-Path -Path $OutputCSVName) -eq "True") {
        if ($ActionFlag -eq "Report") {
            Write-Host Detailed license report is available in: -NoNewline -Foregroundcolor Yellow; Write-Host $OutputCSVName
            Write-Host The report has $ProcessedCount records.
        }
        elseif ($ActionFlag -eq "Mgmt") {
            Write-Host License assignment/removal log file is available in: -NoNewline -Foregroundcolor Yellow; Write-Host $OutputCSVName
        } 
        $Prompt = New-Object -ComObject wscript.shell  
        $UserInput = $Prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)  
        If ($UserInput -eq 6) {  
            Invoke-Item "$OutputCSVName"  
        } 
    }
    else {
        Write-Host No records found
    }
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    Write-Progress -Activity Export CSV -Completed
}

#Get user's details
Function Get_UserInfo {
    $global:DisplayName = $_.DisplayName
    $global:UPN = $_.UserPrincipalName
    $global:Licenses = $_.AssignedLicenses.SkuId
    $SigninStatus = $_.AccountEnabled
    if ($SigninStatus -eq $False) { 
        $global:SigninStatus = "Disabled" 
    }
    else {
        $global:SigninStatus = "Enabled"
    }
    $global:Department = $_.Department
    $global:JobTitle = $_.JobTitle
    if ($Department -eq $null) {
        $global:Department = "-"
    }
    if ($JobTitle -eq $null) {
        $global:JobTitle = "-"
    }
}

Function Get_License_FriendlyName {
    $FriendlyName = @()
    $LicensePlan = @()    
    #Convert license plan to friendly name 
    foreach ($License in $Licenses) {   
        $LicenseItem = $SkuIdHash[$License]  
        $EasyName = $FriendlyNameHash[$LicenseItem]  
        if (!($EasyName)) {
            $NamePrint = $LicenseItem 
        }  
        else {
            $NamePrint = $EasyName 
        } 
        $FriendlyName = $FriendlyName + $NamePrint
        $LicensePlan = $LicensePlan + $LicenseItem
    }
    $global:LicensePlans = $LicensePlan -join ","
    $global:FriendlyNames = $FriendlyName -join ","
}

Function Set_UsageLocation {
    if ($LicenseUsageLocation -ne "") {
        "Assigning Usage Location $LicenseUsageLocation to $UPN" |  Out-File $OutputCSVName -Append
        Update-MgUser -UserId $UPN -UsageLocation $LicenseUsageLocation
        if(!($?))
        {
         "Error occurred while assigning usage location to $UPN or user not found" |  Out-File $OutputCSVName -Append
         Continue
         }
    }
    else {
        "Usage location is mandatory to assign license. Please set Usage location for $UPN" |  Out-File $OutputCSVName -Append
        Continue
    }
}

Function Assign_Licenses {
    "Assigning $LicenseNames license to $UPN" | Out-File $OutputCSVName -Append
    Set-MgUserLicense -UserId $UPN -AddLicenses @{SkuId = $SkuPartNumberHash[$LicenseNames] } -RemoveLicenses @() | Out-Null
    if ($?) {
        "License assigned successfully" | Out-File $OutputCSVName -Append
    }
    else {
        "License assignment failed" | Out-file $OutputCSVName -Append
    }
}

Function Remove_Licenses {
    $SkuPartNumber = @()
    foreach ($Temp in $License) {
        $SkuPartNumber += $SkuIdHash[$Temp]
    }
    $SkuPartNumber = $SkuPartNumber -join (",")
    Write-Progress -Activity "`n     Removing $SkuPartNumber license from $UPN "`n"  Processed users: $ProcessedCount"
    "Removing $SkuPartNumber license from $UPN" | Out-File $OutputCSVName -Append
    Set-MgUserLicense -UserId $UPN -RemoveLicenses @($License) -AddLicenses @() | Out-Null
    if ($?) {
        "License removed successfully" | Out-File $OutputCSVName -Append
    }
    else {
        "License removal failed" | Out-file $OutputCSVName -Append
    }
}

Function main() {
    Disconnect-MgGraph -ErrorAction SilentlyContinue|Out-Null
    Connect_MgGraph
    $Result = ""  
    $Results = @() 
    $FriendlyNameHash = @{}
    Import-Csv -Path .\LicenseFriendlyName.csv -ErrorAction Stop | ForEach-Object {
    $FriendlyNameHash[$_.string_id] = $_.Product_Display_Name
}
    $SkuPartNumberHash = @{} 
    $SkuIdHash = @{} 
    Get-MgSubscribedSku -All | Select-Object SkuPartNumber, SkuId | ForEach-Object {
        $SkuPartNumberHash.add(($_.SkuPartNumber), ($_.SkuId))
        $SkuIdHash.add(($_.SkuId), ($_.SkuPartNumber))
    }

    Do {                 
        if ($Action -eq "") {                       
            Write-Host ""
            Write-host `nOffice 365 License Reporting -ForegroundColor Yellow
            Write-Host  "    1.Get all licensed users" -ForegroundColor Cyan
            Write-Host  "    2.Get all unlicensed users" -ForegroundColor Cyan
            Write-Host  "    3.Get users with specific license type" -ForegroundColor Cyan
            Write-Host  "    4.Get all disabled users with licenses" -ForegroundColor Cyan
            Write-Host  "    5.Office 365 license usage report" -ForegroundColor Cyan
            Write-Host `nOffice 365 License Management -ForegroundColor Yellow
            Write-Host  "    6.Bulk:Assign a license to users (input CSV)" -ForegroundColor Cyan
            Write-Host  "    7.Bulk:Assign multiple licenses to users (input CSV)" -ForegroundColor Cyan
            Write-Host  "    8.Remove all license from a user" -ForegroundColor Cyan
            Write-Host  "    9.Bulk:Remove all licenses from users (input CSV)" -ForegroundColor Cyan
            Write-Host  "    10.Remove specific license from all users" -ForegroundColor Cyan
            Write-Host  "    11.Remove all license from disabled users" -ForegroundColor Cyan
            Write-Host  "    0.Exit" -ForegroundColor Cyan
            Write-Host ""
            $GetAction = Read-Host 'Please choose the action to continue' 
        }
        else {
            $GetAction = $Action
        }

        Switch ($GetAction) {
            1 {
                $OutputCSVName = ".\O365UserLicenseReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
                 $RequiredProperties=@('UserPrincipalName','DisplayName','AccountEnabled','Department','JobTitle','AssignedLicenses') 
                Write-Host Generating licensed users report...
                $ProcessedCount = 0
                Get-MgUser -All -Property $RequiredProperties | Where-Object {($_.AssignedLicenses.Count) -ne 0 } | ForEach-Object {
                    $ProcessedCount++
                    Get_UserInfo
                    Write-Progress -Activity "`n     Processed users count: $ProcessedCount "`n"  Currently Processing: $DisplayName"
                    Get_License_FriendlyName
                    $Result = @{'Display Name' = $Displayname; 'UPN' = $UPN; 'License Plan' = $LicensePlans; 'License Plan Friendly Name' = $FriendlyNames; 'Account Status' = $SigninStatus; 'Department' = $Department; 'Job Title' = $JobTitle }
                    $Results = New-Object PSObject -Property $Result
                    $Results | select-object 'Display Name', 'UPN', 'License Plan', 'License Plan Friendly Name', 'Account Status', 'Department', 'Job Title' | Export-Csv -Path $OutputCSVName -Notype -Append
                }
                $ActionFlag = "Report"
                Open_OutputFile
            }

            2 {
                $OutputCSVName = ".\O365UnlicenedUserReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
                 $RequiredProperties=@('UserPrincipalName','DisplayName','AccountEnabled','Department','JobTitle','AssignedLicenses') 
                Write-Host Generating Unlicensed users report...
                $ProcessedCount = 0
                Get-MgUser -All -Property $RequiredProperties | Where-Object {($_.AssignedLicenses.Count) -eq 0 } | ForEach-Object {
                    $ProcessedCount++
                    Get_UserInfo
                    Write-Progress -Activity "`n     Processed users count: $ProcessedCount "`n"  Currently Processing: $DisplayName"
                    $Result = @{'Display Name' = $Displayname; 'UPN' = $UPN; 'Department' = $Department; 'Signin Status' = $SigninStatus; 'Job Title' = $JobTitle }
                    $Results = New-Object PSObject -Property $Result
                    $Results | select-object 'Display Name', 'UPN', 'Department', 'Job Title', 'Signin Status' | Export-Csv -Path $OutputCSVName -Notype -Append
                }
                $ActionFlag = "Report"
                Open_OutputFile
            }

            3 {
                $OutputCSVName = "./O365UsersWithSpecificLicenseReport__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
                if ($LicenseName -eq "") {
                    $LicenseName = Read-Host "Enter the license SKU(Eg:Enterprisepack)"
                }
                Write-Host Getting users with $LicenseName license...
                $ProcessedCount = 0
                 $RequiredProperties=@('UserPrincipalName','DisplayName','AccountEnabled','Department','JobTitle','AssignedLicenses') 
                if ($SkuPartNumberHash.Keys -icontains $LicenseName) {
                    Get-MgUser -All -Property $RequiredProperties| Where-Object{(($_.AssignedLicenses).SkuId) -eq $SkuPartNumberHash[$LicenseName]} | ForEach-Object {
                        $ProcessedCount++
                        Get_UserInfo
                        Write-Progress -Activity "`n     Processed users count: $ProcessedCount "`n"  Currently Processing: $DisplayName"
                        Get_License_FriendlyName
                        $Result = @{'Display Name' = $Displayname; 'UPN' = $UPN; 'License Plan' = $LicensePlans; 'License Plan_Friendly Name' = $FriendlyNames; 'Account Status' = $SigninStatus; 'Department' = $Department; 'Job Title' = $JobTitle }
                        $Results = New-Object PSObject -Property $Result
                        $Results | select-object 'Display Name', 'UPN', 'License Plan', 'License Plan_Friendly Name', 'Account Status', 'Department', 'Job Title' | Export-Csv -Path $OutputCSVName -Notype -Append
                    }
                }
                else {
                    Write-Host $LicenseName is not used in your organization. Please check the license name or run the License Usage Report to know the licenses in your org -ForegroundColor Red
                }
                #Clearing license name for next iteration
                $LicenseName = ""
                $ActionFlag = "Report"
                Open_OutputFile
            }

            4 {
                $OutputCSVName = "./O365DiabledUsersWithLicense__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
                $ProcessedCount = 0
                 $RequiredProperties=@('UserPrincipalName','DisplayName','AccountEnabled','Department','JobTitle','AssignedLicenses') 
                Write-Host Finding disabled users still licensed in Office 365...
                Get-MgUser -All -Property $RequiredProperties| Where-Object { ($_.AccountEnabled -eq $false) -and (($_.AssignedLicenses).Count -ne 0) } | ForEach-Object {
                    $ProcessedCount++
                    Get_UserInfo
                    Write-Progress -Activity "`n     Processed users count: $ProcessedCount "`n"  Currently Processing: $DisplayName"
                    Get_License_FriendlyName
                    $Result = @{'Display Name' = $Displayname; 'UPN' = $UPN; 'License Plan' = $LicensePlans; 'License Plan_Friendly Name' = $FriendlyNames; 'Department' = $Department; 'Job Title' = $JobTitle }
                    $Results = New-Object PSObject -Property $Result
                    $Results | select-object 'Display Name', 'UPN', 'License Plan', 'License Plan_Friendly Name', 'Department', 'Job Title' | Export-Csv -Path $OutputCSVName -Notype -Append
                }
                $ActionFlag = "Report"
                Open_OutputFile
            }

            5 {
                $OutputCSVName = "./Office365LicenseUsageReport__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
                Write-Host Generating Office 365 license usage report...
                $ProcessedCount = 0
                Get-MgSubscribedSku | ForEach-Object {
                    $ProcessedCount++
                    $AccountSkuID = $_.SkuID
                    $LicensePlan = $_.SkuPartNumber
                    $ActiveUnits = $_.PrepaidUnits.Enabled
                    $ConsumedUnits = $_.ConsumedUnits
                    Write-Progress -Activity "`n     Retrieving license info "`n"  Currently Processing: $LicensePlan"
                    $EasyName = $FriendlyNameHash[$LicensePlan]  
                    if (!($EasyName))  
                    { $FriendlyName = $LicensePlan }  
                    else  
                    { $FriendlyName = $EasyName } 
                    $Result = @{'AccountSkuId' = $AccountSkuID;'AccountSkuPartNumber' = $LicensePlan; 'License Plan_Friendly Name' = $FriendlyName; 'Active Units' = $ActiveUnits; 'Consumed Units' = $ConsumedUnits }
                    $Results = New-Object PSObject -Property $Result
                    $Results | select-object 'AccountSkuId','AccountSkuPartNumber', 'License Plan_Friendly Name', 'Active Units', 'Consumed Units' | Export-Csv -Path $OutputCSVName -Notype -Append
                }
                $ActionFlag = "Report"
                Open_OutputFile
            }

            6 {
                $OutputCSVName = "./Office365LicenseAssignment_Log__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
                $UserNamesFile = Read-Host "Enter the CSV file containing user names(Eg:D:/UserNames.txt)"
                #We have an input file, read it into memory
                $UserNames = @()
                $UserNames = Import-Csv -Header "UPN" $UserNamesFile
                $ProcessedCount = 0
                $LicenseNames = Read-Host "Enter the license name(Eg:Enterprisepack)"
                Write-Host Assigning license to users...
                if ($SkuPartNumberHash.Keys -icontains $LicenseNames) {
                    foreach ($Item in $UserNames) {
                        $ProcessedCount++
                        $UPN = $Item.UPN
                        Write-Progress -Activity "`n     Assigning $LicenseNames license to $UPN "`n"  Processed users: $ProcessedCount"
                        $UsageLocation = (Get-MgUser -UserId $UPN -Property UsageLocation).UsageLocation

                        if ($UsageLocation -eq $null) {
                            Set_UsageLocation
                        }
                       
                        Assign_Licenses
                        
                    }
                }
                else {
                    Write-Host $LicenseNames is not used in your organization. Please check the license name or run the License Usage Report to know the licenses in your org -ForegroundColor Red
                }
                #Clearing license name and input file location for next iteration
                $LicenseNames = ""
                $UserNamesFile = ""
                $ActionFlag = "Mgmt"
                Open_OutputFile
            }

            7 {
                $OutputCSVName = "./Office365LicenseAssignment_Log__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
                $UserNamesFile = Read-Host "Enter the CSV file containing user names(Eg:D:/UserNames.txt)"
                #We have an input file, read it into memory
                $UserNames = @()
                $UserNames = Import-Csv -Header "UPN" $UserNamesFile
                $Flag = ""
                $ProcessedCount = 0
                $License = Read-Host "Enter the license names(Eg:LicensePlan1,LicensePlan2)"
                $License = $License.Replace(' ', '')
                $License = $License.split(",")
                foreach ($LicenseName in $License) {  
                    if ($SkuPartNumberHash.Keys -inotcontains $LicenseName) {
                        $Flag = "Terminate"
                        Write-Host $LicenseName is not used in your organization. Please check the license name or run the License Usage Report to know the licenses in your org -ForegroundColor Red
                    }
                }
                if ($Flag -eq "Terminate") {
                    Write-Host Please re-run the script with appropriate license name -ForegroundColor Yellow
                }
                else {
                    Write-Host Assigning licenses to Office 365 users...
                    foreach ($Item in $UserNames) {
                        $UPN = $Item.UPN
                        $ProcessedCount++
                        $UsageLocation = (Get-MgUser -UserId $UPN -Property UsageLocation).UsageLocation
                        if ($UsageLocation -eq $null) {
                            Set_UsageLocation
                        }
                       
                        Write-Progress -Activity "`n     Assigning licenses to $UPN "`n"  Processed users: $ProcessedCount"
                        foreach ($LicenseNames in $License) {
                             Assign_Licenses
                        }
                        
                    }
                }
                #Clearing license names and input file location for next iteration
                $LicenseNames = ""
                $UserNamesFile = ""
                $ActionFlag = "Mgmt"
                Open_OutputFile
            }
       

            8 {
                $Identity = Read-Host Enter User UPN
                $UserInfo = Get-MgUser -UserId $Identity -Property "DisplayName,AssignedLicenses"
                #Checking whether the user is available
                if ($UserInfo -eq $null) {
                    Write-Host User $Identity does not exist. Please check the user name. -ForegroundColor Red
                }
                else {
                    $Licenses = $UserInfo.AssignedLicenses.SkuId
                    $SkuPartNumber = @()
                    if ($Licenses.count -eq 0) {
                        Write-Host No license assigned to the user $Identity. 
                    }
                    else {
                        foreach ($Temp in $Licenses) {
                            $SkuPartNumber += $SkuIdHash[$Temp]
                        }
                        $SkuPartNumber = $SkuPartNumber -join (",")
                        Write-Host Removing $SkuPartNumber license from $Identity
                        Set-MgUserLicense -UserId $Identity -RemoveLicenses @($Licenses) -AddLicenses @() | Out-Null
                        Write-Host Action completed -ForegroundColor Green                        
                    }
                }  
            }

            9 {
                $OutputCSVName = "./Office365LicenseRemoval_Log__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
                $UserNamesFile = Read-Host "Enter the CSV file containing user names(Eg:D:/UserNames.txt)"
                #We have an input file, read it into memory
                $UserNames = @()
                $UserNames = Import-Csv -Header "UPN" $UserNamesFile
                $ProcessedCount = 0
                foreach ($Item in $UserNames) {
                    $UPN = $Item.UPN
                    $ProcessedCount++
                    $License = (Get-MgUser -UserId $UPN -Property AssignedLicenses).AssignedLicenses.SkuId
                    if ($License.count -eq 0) {
                        "No License Assigned to this user $UPN" | Out-File $OutputCSVName -Append
                    }
                    else {
                        Remove_Licenses
                    }
                }
                $ActionFlag = "Mgmt"
                Open_OutputFile 
            }

            10 {
                $OutputCSVName = "./O365LicenseRemoval_Log__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
                $Licenses = Read-Host "Enter the license name(Eg:LicensePlan)"
                $License = $SkuPartNumberHash[$Licenses]
                $ProcessedCount = 0
                if ($SkuPartNumberHash.Values -icontains $License) {
                    Get-MgUser -All -Property UserPrincipalName,AssignedLicenses | Where-Object { ($_.AssignedLicenses).SkuId -eq $License } | ForEach-Object {
                        $ProcessedCount++
                        $UPN = $_.UserPrincipalName
                        Remove_Licenses
                    }
                }
                else {
                    Write-Host $License not used in your organization. Please check the license name or run the License Usage Report to know the licenses in your org -ForegroundColor Red
                }
                $ActionFlag = "Mgmt"
                Open_OutputFile
            }  

            11 {
                $OutputCSVName = "./O365LicenseRemoval_Log__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
                Write-Host Removing license from disabled users...
                $ProcessedCount = 0
                Get-MgUser -All -Property UserPrincipalName,AssignedLicenses,AccountEnabled | Where-Object { ($_.AccountEnabled -eq $false) -and (($_.AssignedLicenses).Count -ne 0) } | ForEach-Object {
                    $ProcessedCount++
                    $UPN = $_.UserPrincipalName
                    $License = $_.AssignedLicenses.SkuId
                    Remove_Licenses
                }
                $ActionFlag = "Mgmt"
                Open_OutputFile
            } 
        }
        if ($Action -ne "") {
            exit 
        }
        if ($MultipleActionsMode.ispresent) {                          
            Start-Sleep -Seconds 2
        } 
        else {
            Exit
        }
    }
    While ($GetAction -ne 0)
    Disconnect-MgGraph
    Write-Host "Disconnected active Microsoft Graph session"
    Clear-Host
}
. main