<#
=============================================================================================
Name:     Get All Enterprise Applications with Their Permissions       
Version:  1.0       
Website:  blog.admindroid.com       


Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Exports all enterprise apps along with its API permissions in Microsoft Entra.
2. The script installs MS Graph PowerShell SDK (if not installed already) upon your confirmation. 
3. Allows to filter applications with specific permissions (eg.,"User.Read.All") assigned.
    -> Admin consented app permissions 
    -> Admin consented delegated permissions
    -> User consented permissions 
4. Fetches the list of ownerless applications in Microsoft Entra. 
5. Find Entra app permissions granted thorough user consent and admin consent. 
6. Generates report for sign-in enabled and disabled applications. 
7. Assists in filtering based on following properties:
    -> Application Name
    -> Application Id
    -> Object Id
    -> API Name
8. Filters apps that are restricted to specific users and accessible to all users. 
9. Lists applications that are hidden and visible to all users in the organization.  
10. Assists in filtering home tenant and external tenant applications.
11. Allows to retrieve enterprise apps with no permissions too.  
12. Exports the result to CSV.
13. The script can be executed with an MFA enabled account too.  
14. It can be executed with certificate-based authentication (CBA) too.
15. The script is schedular-friendly.   

For detailed Script execution: https://blog.admindroid.com/export-all-enterprise-apps-and-their-assigned-permission-in-microsoft-entra/


============================================================================================
#>

Param (
    [Parameter(Mandatory = $false)]
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbPrint,
    [string]$ApplicationId,
    [string]$ApplicationName,
    [string]$ObjectId,
    [string]$APIName,
    [ValidateSet("VisibleApps","HiddenApps")]
    [string]$AppVisibility,
    [ValidateSet("HomeTenant","ExternalTenant")]
    [string]$AppOrigin,
    [ValidateSet("Enabled", "Disabled")]
    [string]$UsersSignIn,
    [ValidateSet("AdminConsent", "UserConsent")]
    [string]$ConsentType,
    [string[]]$AdminConsentApplicationPermissions,
    [string[]]$AdminConsentDelegatedPermissions,
    [string[]]$UserConsents,
    [switch]$AccessScopeToAllUsers,
    [switch]$RoleAssignmentRequiredApps,
    [switch]$OwnerlessApps,
    [switch]$IncludeAppsWithNoPermissions
)

function Connect_MgGraph {
    $MsGraphModule = Get-Module Microsoft.Graph -ListAvailable
    if ($MsGraphModule -eq $null) {
        Write-Host "`nImportant: Microsoft Graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."
        $confirm = Read-Host "Are you sure you want to install Microsoft Graph module? [Y] Yes [N] No"
        if ($confirm -match "[yY]") {
            Write-Host "Installing Microsoft Graph Module..."
            Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber -Force
        }
        else {
            Write-Host "Microsoft Graph PowerShell module is required to run this script. Please install module using 'Install-Module Microsoft.Graph' cmdlet."
            Exit
        }
    }

    if ($CreateSession.IsPresent) {
        Disconnect-MgGraph
    }

    Write-Host "`nConnecting to Microsoft Graph..."
    if (($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbPrint -ne "")) {
        Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbPrint -NoWelcome
    }
    else {
        Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome
    }
}

Connect_MgGraph

$ExportCSV = "$(Get-Location)\EnterpriseApps_and_their_Permissions_report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$TenantGUID= (Get-MgOrganization).Id
Write-Host "`nRetreiving the Enterprise applications with admin consents and user consents"
$AppCount = 0 
$PrintCount = 0

Get-MgServicePrincipal -All | ForEach-Object {
    $Print = 1
    $AppCount++
    $ServicePrincipalType = $_.ServicePrincipalType
    $AppName = $_.DisplayName
    Write-Progress -Activity "Processed Enterprise apps: $($AppCount) $($AppName)" 
    $AppId = $_.AppId
    $ObjId  = $_.Id
    $CreatedDateTime = [datetime]@($_.AdditionalProperties.Values)[0]
    $AccountEnabled = if ($_.AccountEnabled) { "Enabled" } else { "Disabled" }
    $Owners = (Get-MgServicePrincipalOwner -ServicePrincipalId $_.Id | ForEach-Object { $_.AdditionalProperties["displayName"] }) -join ", "
    $Tags = $_.Tags
    $IsRoleAssignmentRequired = $_.AppRoleAssignmentRequired

    if (-not $Owners) { $Owners = "-" }
    if ($Tags -contains "HideApp") { $UserVisibility="Hidden" }
    else { $UserVisibility="Visible" }
    if ($IsRoleAssignmentRequired -eq $true){ $AccessScope="Only assigned users can access" }
    else { $AccessScope="All users can access" }
    $AppOwnerOrgId=$_.AppOwnerOrganizationId
    if ($AppOwnerOrgId -eq $TenantGUID){ $AppOrg="Home tenant" }
    else { $AppOrg="External tenant" }
    
    if (($ApplicationId.Length -ne 0) -and ($ApplicationId -ne $AppId)) { $Print = 0 }
    if (($ApplicationName.Length -ne 0) -and ($ApplicationName -ne $AppName)) { $Print = 0 }
    if (($ObjectId.Length -ne 0) -and ($ObjectId -ne $ObjId)) { $Print = 0 }
    if ($UsersSignIn -eq "Enabled" -and $_.AccountEnabled -ne $true) { $Print = 0 }
    if ($UsersSignIn -eq "Disabled" -and $_.AccountEnabled -ne $false) { $Print = 0 }
    if (($AppVisibility -eq "VisibleApps") -and ($UserVisibility -ne "Visible")){ $Print=0 }
    if (($AppVisibility -eq "HiddenApps") -and ($UserVisibility -ne "Hidden")){ $Print=0 }
    if (($AccessScopeToAllUsers.IsPresent) -and ($AccessScope -eq "Only assigned users can access")){ $Print=0 }
    if (($RoleAssignmentRequiredApps.IsPresent) -and ($AccessScope -eq "All users can access")){ $Print=0 }
    if (($OwnerlessApps.IsPresent) -and ($Owners -ne "-")){ $Print=0 }
    if ($AppOrigin -eq "HomeTenant" -and ($AppOrg -eq "External tenant")){ $Print=0 }
    if ($AppOrigin -eq "ExternalTenant" -and ($AppOrg -eq "Home tenant")){ $Print=0 }
    
    $DelegatedGrants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $ObjId -All 
    $AppAssignments  = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ObjId -All
    $AllAPIids = @($DelegatedGrants.ResourceId; $AppAssignments.ResourceId) | Sort-Object -Unique

    if (-not $AllAPIids) { $AllAPIids = @('-') }  

    foreach ($ResourceId in $AllAPIids) {
        if ($ResourceId -eq '-') {
            $ResourceName = '-'
            $AdminDelegated = '-'
            $UserDelegated  = '-'
            $AdminApps      = '-'
        }
        else {
            $ResourceSp = Get-MgServicePrincipal -ServicePrincipalId $ResourceId
            $ResourceName = $ResourceSp.DisplayName

            $AdminDelegated = $DelegatedGrants | Where-Object { $_.ResourceId -eq $ResourceId -and $_.ConsentType -eq "AllPrincipals" } | ForEach-Object{ $_.Scope.Trim()}
            $AdminDelegated = if(-not $AdminDelegated) {"-"} else {$AdminDelegated -split "\s+" -join ", "}

            $UserDelegated  = $DelegatedGrants | Where-Object { $_.ResourceId -eq $ResourceId -and $_.ConsentType -eq "Principal" } | ForEach-Object {$_.Scope.Trim()}
            $UserDelegated = if(-not $UserDelegated){"-"} else {$UserDelegated -split "\s+" -join ", "}

            $AdminApps = $AppAssignments | Where-Object { $_.ResourceId -eq $ResourceId } |
                         ForEach-Object {
                            $role = $ResourceSp.AppRoles | Where-Object Id -eq $_.AppRoleId
                            if ($role) { $role.Value }
                         } 
            $AdminApps = if(-not $AdminApps){"-"} else {$AdminApps -join ", "}
        }
       
        if ((-not $IncludeAppsWithNoPermissions.IsPresent) -and ($AdminDelegated[0] -eq "-" -and $AdminApps[0] -eq "-" -and $UserDelegated[0] -eq "-")) { $Print = 0 }
        if ($AdminConsentApplicationPermissions -and ((($AdminApps -split ", ") | Where-Object { $_ -in $AdminConsentApplicationPermissions }).Count -eq 0)) { $Print = 0 }
        if ($AdminConsentDelegatedPermissions -and ((($AdminDelegated -split ", ") | Where-Object { $_ -in $AdminConsentDelegatedPermissions}).Count -eq 0)) { $Print = 0 }
        if ($UserConsents -and ((($UserDelegated -split ", ") | Where-Object {$_ -in $UserConsents}).Count -eq 0)) { $Print = 0 }
        if (($APIName.Length -ne 0) -and ($APIName -ne $ResourceName)) { $Print = 0 }
        if ($ConsentType -eq "AdminConsent" -and $UserDelegated -ne '-') { $Print = 0 }
        if ($ConsentType -eq "UserConsent" -and ($AdminApps -ne '-' -or $AdminDelegated -ne '-')) { $Print = 0 }

        if ($Print -eq 1){
            $PrintCount++
            [PSCustomObject]@{
                'App Name' = $AppName
                'Object Id'= $ObjId
                'API Name' = $ResourceName
                'Admin Consented App Permissions' = $AdminApps
                'Admin Consented Delegated Permissions' = $AdminDelegated
                'User Consented Permissions' = $UserDelegated
                'Owners'   = $Owners
                'Users Sign In' = $AccountEnabled
                'User Visibility'= $UserVisibility
                'Role Assignment Required'= $AccessScope
                'Service Principal Type'= $ServicePrincipalType
                'App Id'   = $AppId
                'App Origin'= $AppOrg
                'App Org Id'= $AppOwnerOrgId
                'API Id'   = $ResourceId
                'Created Date'  = $CreatedDateTime
            } | Export-Csv -Path $ExportCSV -Append -NoTypeInformation
        }
    }
}

Disconnect-MgGraph | Out-Null

Write-Host `nScript executed successfully.
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green `n`n

if(((Test-Path -Path $ExportCSV) -eq "True"))
{
    Write-Host `nThe script processed $AppCount enterprise apps and the output file contains $PrintCount records.
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6)
    {  
        Invoke-Item "$ExportCSV"
    }
    Write-Host "The generated report is available in: " -NoNewline -ForegroundColor Yellow; Write-Host "$($ExportCSV)"
}
else
{
    Write-Host "No user found" -ForegroundColor Red
}