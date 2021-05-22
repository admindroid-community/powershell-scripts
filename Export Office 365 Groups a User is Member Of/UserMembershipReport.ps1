<#
=============================================================================================
Name:           Office365 User Membership Report Using PowerShell
Description:    This script exports Office 365 user's group details to CSV
Version:        1.0
Website:        o365reports.com
Script by:      O365Reports Team
For detailed script execution:https://o365reports.com/2021/04/15/export-office-365-groups-a-user-is-member-of-using-powershell/
============================================================================================
#>
param
(
    [string] $UserName = $null,
    [string] $Password = $null,
    [String] $UsersIdentityFile,
    [Switch] $GuestUsersOnly,
    [Switch] $DisabledUsersOnly,
    [Switch] $UsersNotinAnyGroup
)

Function UserDetails {
    if ([string]$UsersIdentityFile -ne "") {
        
        $IdentityList = Import-Csv -Header "UserIdentityValue" $UsersIdentityFile
        foreach ($IdentityValue in $IdentityList) {
            $CurIdentity = $IdentityValue.UserIdentityValue
            try {
                $LiveUser = Get-AzureADUser -ObjectId "$CurIdentity" -ErrorAction SilentlyContinue
                if ($GuestUsersOnly.IsPresent -and $LiveUser.UserType -ne "Guest") {
                    continue
                }
                if ($DisabledUsersOnly.IsPresent -and $LiveUser.AccountEnabled -eq $true) {
                    continue
                }
                ProcessUser
            }
            catch {
                Write-Host Given UserIdentity: $CurIdentity is not valid/found.
            }
        }
    }
    else {
        if ($GuestUsersOnly.Ispresent -and $DisabledUsersOnly.Ispresent) {
            Get-AzureADUser -Filter "UserType eq 'Guest'" | Where-Object { $_.AccountEnabled -eq $false } | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }
        }
        elseif ($DisabledUsersOnly.Ispresent) {
            Get-AzureADUser | Where-Object { $_.AccountEnabled -eq $false } | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }  
        }
        elseif ($GuestUsersOnly.Ispresent) {
            Get-AzureADUser -Filter "UserType eq 'Guest'" | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }
        }
        else {
            Get-AzureADUser -All:$true | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }
        }
    }
}
Function ProcessUser {
    $GroupList = @()
    $Roles = @()
    $global:ProcessedUsers = $global:ProcessedUsers + 1
    $UPN = $LiveUser.UserPrincipalName
    $ObjectId = $LiveUser.ObjectId.ToString()
    $Name = $LiveUser.DisplayName
    Write-Progress -Activity "Processing $Name" -Status "Processed Users Count: $global:ProcessedUsers" 
    $UserMembership = Get-AzureADUserMembership -ObjectId $ObjectId
    
    $AllGroupData = $UserMembership | Where-object { $_.ObjectType -eq "Group" }
    if ($null -eq $AllGroupData) {
        $GroupName = " - "
    }
    else {
        if ($UsersNotinAnyGroup.IsPresent) {
            return
        }
        $AllGroupData | Select-Object DisplayName | ForEach-Object {
            $GroupList += $_.DisplayName -join ","
            $GroupName = ($GroupList -join ",")
            }        
    }
    $AllRoles = $UserMembership | Where-object { $_.ObjectType -eq "Role" } | Select-Object DisplayName
    if ($null -eq $AllRoles) { 
        $RolesList = " - " 
    }
    else {
        $AllRoles | ForEach-Object {
            $Roles += $_.DisplayName -join ","
            $RolesList = ($Roles -join ",")
        }
    }
    if ($LiveUser.AccountEnabled -eq $True) {
        $AccountStatus = "Enabled"
    }
    else {
        $AccountStatus = "Disabled"
    }
    if ($null -eq $LiveUser.Department) {
        $Department = " - " 
    }
    else {
        $Department = $LiveUser.Department
    }
    if ($LiveUser.AssignedLicenses -ne "") { 
        $LicenseStatus = "Licensed" 
    }
    else {
        $LicenseStatus = "Unlicensed" 
    }
    ExportResults
}
Function ExportResults {
    $global:ExportedUsers = $global:ExportedUsers + 1
    $ExportResult = @{'Display Name' = $Name; 'Email Address' = $UPN; 'Group Name(s)' = $GroupName; 'License Status' = $LicenseStatus; 'Account Status' = $AccountStatus;'Department' = $Department;'Admin Roles' = $RolesList }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-Object  'Display Name', 'Email Address', 'Group Name(s)', 'License Status', 'Account Status', 'Department', 'Admin Roles'  | Export-csv -path $ExportCSVFileName -NoType -Append    
}
Function Connection {
    $AzureAd = (Get-Module AzureAD -ListAvailable).Name
    if ($Empty -eq $AzureAd) {
        Write-host "Mandatory module is unavailable: Install AzureAd module to run the script successfully. Please choose (Y or y) to say Yes" 
        $confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
        if ($confirm -match "[yY]") { 
            Write-host "Installing AzureAD"
            Install-Module AzureAd -Allowclobber -Repository PSGallery -Force
            Write-host "Required Module is installed in the machine Successfully"
        }
        else { 
            Write-host "Exiting. `nNote:The Module AzureAD must be available in your system to run the script" 
            Exit 
        }
    }
    #Importing Module by default will avoid the cmdlet unrecognized error 
    Import-Module AzureAd -ErrorAction SilentlyContinue -Force
    Write-Host "Connecting to AzureAD..." 
    #Storing credential in script for scheduling purpose/Passing credential as parameter   
    if (($UserName -ne "") -and ($Password -ne "")) {   
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force   
        $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword   
        Connect-AzureAD -Credential $Credential | Out-Null
    }   
    else {   
        Connect-AzureAD | Out-Null
    }
}

Connection
$global:ProcessedUsers = 0
$global:ExportedUsers = 0
$ExportCSVFileName = ".\UserMembershipReport_$((Get-Date -format MMM-dd` hh-mm-ss` tt).ToString()).csv"
UserDetails
#Open output file after execution
if ((Test-Path -Path $ExportCSVFileName) -eq "True") { 
    Write-Progress -Activity "--" -Completed
    Write-Host "The Output result has $global:ExportedUsers users details"
    Write-Host `nThe Output file available in $ExportCSVFileName -ForegroundColor Green 
    $prompt = New-Object -ComObject wscript.shell    
    $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
    If ($userInput -eq 6) {    
        Invoke-Item "$ExportCSVFileName"
    }  
}
else {
    Write-Host `nNo data/user found with the specified criteria
}