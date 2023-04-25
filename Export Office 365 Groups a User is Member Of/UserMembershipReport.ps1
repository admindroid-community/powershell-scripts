<#
=============================================================================================
Name:           Office365 User Membership Report Using PowerShell
Description:    This script exports Office 365 user's group details to CSV
Version:        3.0
Website:        o365reports.com
For detailed script execution:https://o365reports.com/2021/04/15/export-office-365-groups-a-user-is-member-of-using-powershell
============================================================================================
#>

param
(
    [String] $UsersIdentityFile,
    [Switch] $GuestUsersOnly,
    [Switch] $DisabledUsersOnly,
    [Switch] $UsersNotinAnyGroup,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint
)
$MsGraphModule =  Get-Module Microsoft.Graph -ListAvailable
if($MsGraphModule -eq $null)
{ 
    Write-host "Important: Microsoft graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
    $confirm = Read-Host Are you sure you want to install Microsoft graph module? [Y] Yes [N] No  
    if($confirm -match "[yY]") 
    { 
        Write-host "Installing Microsoft graph module..."
        Install-Module Microsoft.Graph -Scope CurrentUser
        Write-host "Microsoft graph module is installed in the machine successfully" -ForegroundColor Magenta 
    } 
    else
    { 
        Write-host "Exiting. `nNote: Microsoft graph module must be available in your system to run the script" -ForegroundColor Red
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
        Write-Host "$ConnectionError" -Foregroundcolor Red
        Exit
    }
}
Write-Host "Microsoft Graph Powershell module is connected successfully" -ForegroundColor Green
Select-MgProfile beta
Function UserDetails {
    if ([string]$UsersIdentityFile -ne "")
    {
        $IdentityList = Import-Csv -Header "UserIdentityValue" $UsersIdentityFile
        foreach ($IdentityValue in $IdentityList) 
        {
            $CurIdentity = $IdentityValue.UserIdentityValue
            try 
            {
                $LiveUser = Get-MgUser -UserId "$CurIdentity" -ExpandProperty MemberOf -ErrorAction SilentlyContinue
                if($GuestUsersOnly.IsPresent -and $LiveUser.UserType -ne "Guest") 
                {
                    continue
                }
                if($DisabledUsersOnly.IsPresent -and $LiveUser.AccountEnabled -eq $true)
                {
                    continue
                }
                ProcessUser
            }
            catch 
            {
                Write-Host Given UserIdentity: $CurIdentity is not valid/found.
            }
        }
    }
    else 
    {
        if ($GuestUsersOnly.Ispresent -and $DisabledUsersOnly.Ispresent) 
        {
            Get-MgUser  -Filter "UserType eq 'Guest'" -ExpandProperty MemberOf -All| Where-Object { $_.AccountEnabled -eq $false } | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }
        }
        elseif ($DisabledUsersOnly.Ispresent) 
        {
            Get-MgUser -ExpandProperty MemberOf -All| Where-Object { $_.AccountEnabled -eq $false } | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }  
        }
        elseif ($GuestUsersOnly.Ispresent) 
        {
            Get-MgUser  -Filter "UserType eq 'Guest'" -ExpandProperty MemberOf -All| ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }
        }
        else 
        {
            Get-MgUser -ExpandProperty MemberOf -All | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }
        }
    }
}
Function ProcessUser {
    $GroupList = @()
    $RolesList = @()
    $Script:ProcessedUsers += 1
    $Name = $LiveUser.DisplayName
    Write-Progress -Activity "Processing $Name" -Status "Processed Users Count: $Script:ProcessedUsers" 
    $UserMembership = Get-MgUserMemberOf -UserId $LiveUser.UserPrincipalName |Select-Object -ExpandProperty AdditionalProperties
    $AllGroupData = $UserMembership | Where-object { $_.'@odata.type'  -eq "#microsoft.graph.group" }
    if ($AllGroupData -eq $null) 
    {
        $GroupName = " - "
    }
    else 
    {
        if ($UsersNotinAnyGroup.IsPresent) 
        {
            return
        }
        $GroupName = (@($AllGroupData.displayName) -join ',') 
    }
    $AllRoles = $UserMembership | Where-object { $_.'@odata.type' -eq "#microsoft.graph.directoryRole" }
    if ($AllRoles -eq $null) { 
        $RolesList = " - " 
    }
    else
    {
        $RolesList = @($AllRoles.displayName) -join ','
    }
    if ($LiveUser.AccountEnabled -eq $True) 
    {
        $AccountStatus = "Enabled"
    }
    else 
    {
        $AccountStatus = "Disabled"
    }
    if ($LiveUser.Department -eq $null) 
    {
        $Department = " - " 
    }
    else 
    {
        $Department = $LiveUser.Department
    }
    if ($LiveUser.AssignedLicenses -ne "")
    { 
        $LicenseStatus = "Licensed" 
    }
    else 
    {
        $LicenseStatus = "Unlicensed" 
    }
    ExportResults
}
Function ExportResults {
    $Script:ExportedUsers += 1
    $ExportResult = [PSCustomObject] @{'Display Name' = $Name; 'Email Address' = $LiveUser.UserPrincipalName; 'Group Name(s)' = $GroupName; 'License Status' = $LicenseStatus; 'Account Status' = $AccountStatus;'Department' = $Department;'Admin Roles' = $RolesList }
    $ExportResult | Export-csv -path $ExportCSVFileName -NoType -Append    
}
$ProcessedUsers = 0
$ExportedUsers = 0
$ExportCSVFileName = ".\UserMembershipReport_$((Get-Date -format MMM-dd` hh-mm-ss` tt).ToString()).csv"
UserDetails
#Open output file after execution
if ((Test-Path -Path $ExportCSVFileName) -eq "True") { 
    Write-Progress -Activity "--" -Completed
    Write-Host "`nThe Output result has " -NoNewline
    Write-Host $Script:ExportedUsers Users -ForegroundColor Magenta -NoNewline
    Write-Host " details"
    Write-Host `nThe Output file available in $ExportCSVFileName -ForegroundColor Green 
    $prompt = New-Object -ComObject wscript.shell    
    $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
    if ($userInput -eq 6) {    
        Invoke-Item "$ExportCSVFileName"
    }  
}
else 
{
    Write-Host `nNo data/user found with the specified criteria -ForegroundColor Red
}