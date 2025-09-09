<#
=============================================================================================
Name:           Find Groups with Disabled Users in Microsoft 365 using PowerShell 
Version:        1.0
website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~

1. Exports disabled users based on group types, such as Microsoft 365, Security, Mail-enabled security, and Distribution list. 
2. Provides counts for disabled members and disabled owners in each group. 
3. Automatically installs the required Microsoft Graph PowerShell module with your confirmation. 
4. The script can be executed with an MFA-enabled account too. 
5. Supports Certificate-based Authentication too. 
6. The script is scheduler friendly. 
7. Exports report results into a CSV file. 

For detailed Script execution:  https://o365reports.com/2025/01/21/find-groups-with-disabled-users-in-microsoft-365/
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$M365GroupsOnly,
    [switch]$SecurityGroupsOnly,
    [switch]$MailEnabledSecurityGroupsOnly,
    [switch]$DistributionListsOnly,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)


$CSVFilePath = "$(Get-Location)\GroupsWithDisabledUsers_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"


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
            Write-Host "Microsoft Graph module is installed in the machine successfully" 
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
        Connect-MgGraph -Scopes "Group.Read.All", "GroupMember.Read.All" -NoWelcome 
    }

    # Verify connection
    if ((Get-MgContext) -ne $null) {
        Write-Host "Connected to Microsoft Graph PowerShell using account: $((Get-MgContext).Account)" 
    } else {
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
}


# Export function to output data to CSV
function ExportResult {
    $ExportResult = @{
        'CreatedDateTime' = $CreatedDateTime
        'Group Name' = $GroupName
        'Group Email Address' = $GroupEmailAddress
        'Group Id' = $GroupId
        'Group Type' = $GroupType
        'Is Dynamic Group' = $IsDynamicGroup
        'Disabled Members' = $DisabledMembers -join(", ")
        'Disabled Members Count' = $DisabledMembersCount
        'Disabled Owners' = $DisabledOwners -join(", ")
        'Disabled Owners Count' = $DisabledOwnersCount
        'Total Members Count' = $GroupMembers.Count
        'Total Owners Count' = $GroupOwners.Count
    }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-Object 'Group Name', 'Group Email Address', 'Group Type', 'Is Dynamic Group', 'Total Members Count', 'Total Owners Count', 'Disabled Members Count', 'Disabled Owners Count', 'Disabled Members', 'Disabled Owners', 'CreatedDateTime', 'Group Id' | Export-Csv -Path $CSVFilePath -NoTypeInformation -Append
}


# Connecting to the Microsoft Graph PowerShell Module
Connect_ToMgGraph

$GroupsWithDisabledUsersCount = 0
$ProcessedGroupsCount = 0

# Get all groups in Microsoft 365 
Get-MgGroup -All | ForEach-Object {    
    $CreatedDateTime = $_.CreatedDateTime
    $GroupName = $_.DisplayName
    $GroupEmailAddress = $_.Mail
    $GroupId = $_.Id
    $GroupTypes = $_.GroupTypes
    $IsDynamicGroup = $false
    
    $ProcessedGroupsCount++
    Write-Progress -Activity "`n     Processed group count: $ProcessedGroupsCount"`n"  Checking members of: $GroupName"

    # Get the group type based on properties
    if ($GroupTypes -contains "Unified") {
        $GroupType = "Microsoft365"
    }
    elseif ($_.Mail -ne $null -and $_.SecurityEnabled -eq $false) {
        $GroupType = "DistributionList"
    }
    elseif ($_.Mail -ne $null -and $_.SecurityEnabled -eq $true) {
        $GroupType = "MailEnabledSecurity"
    } 
    else {
        $GroupType = "Security"
    }

    # check whether it is a dynamic membership group
    if ($GroupTypes -contains "DynamicMembership") {
        $IsDynamicGroup = $true
    }

    # If group email address is empty
    if($GroupEmailAddress -eq $null) {
        $GroupEmailAddress = "-"
    }

    # Apply filters based on switches
    if ($SecurityGroupsOnly -and $GroupType -ne "Security") { return }
    if ($DistributionListsOnly -and $GroupType -ne "DistributionList") { return }
    if ($MailEnabledSecurityGroupsOnly -and $GroupType -ne "MailEnabledSecurity") { return }
    if ($M365GroupsOnly -and $GroupType -ne "Microsoft365") { return }


    $GroupMembers = Get-MgGroupMemberAsUser -GroupId $GroupId -All -Property AccountEnabled, UserPrincipalName 
    $GroupOwners = Get-MgGroupOwnerAsUser -GroupId $GroupId -All -Property AccountEnabled, UserPrincipalName 

    $DisabledMembers = $GroupMembers | Where-Object { $_.AccountEnabled -eq $false } | Select-Object -ExpandProperty UserPrincipalName
    $DisabledOwners = $GroupOwners | Where-Object { $_.AccountEnabled -eq $false } | Select-Object -ExpandProperty UserPrincipalName
    $DisabledMembersCount = $DisabledMembers.Count
    $DisabledOwnersCount = $DisabledOwners.Count

    # Export results if there are disabled users
    if (($DisabledMembersCount -gt 0) -or ($DisabledOwnersCount -gt 0)) {
        $GroupsWithDisabledUsersCount++
        if ($DisabledMembers -eq $null) { $DisabledMembers = "-" }
        if ($DisabledOwners -eq $null) { $DisabledOwners = "-" }
        ExportResult
    }
}


# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1900+ Microsoft 365 reports. ~~" -ForegroundColor Green `n

if((Test-Path -Path $CSVFilePath) -eq "True") {   
    Write-Host  " There are $GroupsWithDisabledUsersCount groups with disabled users."  
    Write-Host  " The Output file is available in: " -NoNewline -ForegroundColor Yellow; Write-Host "$CSVFilePath"  
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open the Output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6) {
        Invoke-Item "$CSVFilePath"
    }
}
else{
    Write-Host "No groups found with Disabled Users for the given criteria." 
}