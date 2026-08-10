$Shell = $Host.UI.RawUI
$Shell.WindowTitle="SysadminGeek"

$Script:GraphScopes = @(
    "User.Read.All"
)
$Script:EntraScopes = @(
    "User.Read.All"
)
$Script:CertificateValidityPeriod = (Get-Date).AddYears(1)

if (-not $Script:ConnectedTenantName) {
    $Script:ConnectedTenantName = ""
}

Function UpdateModules {
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

    if((($PSVersionTable::PSVersion.Major) -ge 7) -and (Get-InstalledModule -Name SharePointPnPPowerShellOnline -ErrorAction:SilentlyContinue))
    {
        # Uninstalling legacy module
        Write-Host Removing SharePointPnPPowerShellOnline. Legacy module Powershell 7. -ForegroundColor Green
        Uninstall-Module SharePointPnPPowerShellOnline -Force -AllVersions -ErrorAction:SilentlyContinue
    }

    if(Get-InstalledModule -Name AzureAD -ErrorAction:SilentlyContinue)
    {
        # Uninstalling legacy module
        Write-Host Removing AzureAD. -ForegroundColor Green
        Uninstall-Module AzureAD -Force -AllVersions -ErrorAction:SilentlyContinue
    }

    if(Get-InstalledModule -Name MSOnline -ErrorAction:SilentlyContinue)
    {
        # Uninstalling legacy module
        Write-Host Removing MSOnline. -ForegroundColor Green
        Uninstall-Module MSOnline -Force -AllVersions -ErrorAction:SilentlyContinue
    }

    $lockModules = @(
        [pscustomobject]@{Name='ExchangeOnlineManagement';keepVersion='3.9.0'} # bug certificate auth for 3.10.0 for secandcompcenter Powershell 5 en 7 AND bug with MFA auth for 3.10.1 for secandcompcenter Powershell 7
        [pscustomobject]@{Name='Microsoft.Entra'; keepVersion='1.2.0'} # 1.3.0 still has login issues with credentials/WAM for other tenants
        [pscustomobject]@{Name='Microsoft.Graph'; keepVersion='2.33.0'} # 2.39.0 still has login issues with credentials/WAM for other tenants
        [pscustomobject]@{Name='Microsoft.Graph.Beta'; keepVersion='2.33.0'} # 2.39.0 still has login issues with credentials/WAM for other tenants
    )

    $modules = Get-InstalledModule
    foreach ( $lockModule in $lockModules ) {
        if ($modules.Name -notcontains $lockModule.Name) {
            write-host Locked module $lockModule.Name not yet installed -ForegroundColor Cyan
            write-host Installing older module $lockModule.Name RequiredVersion $lockModule.keepVersion -ForegroundColor Green

            # 1 Install desired version
            Install-Module $lockModule.Name -RequiredVersion $lockModule.keepVersion -AllowClobber -Scope CurrentUser
        }
    }

    foreach ($module in $modules) {
        write-host Checking update for $module.Name -ForegroundColor Cyan

        $versionLocked = $False

        foreach ( $lockModule in $lockModules ) {
            if ( ( $module.Name.StartsWith( $lockModule.Name ) ) -and ( -not $module.Name.StartsWith( "$($lockModule.Name).Beta" ) ) ) {
                $versionLocked = $True

                if ( $module.Name -eq $lockModule.Name ) {
                    write-host Installing older module $lockModule.Name RequiredVersion $lockModule.keepVersion and uninstalling all other versions -ForegroundColor Green

                    # 1 Installeer gewenste versie
                    if ( $module.Version -ne $lockModule.keepVersion ) {
                        Install-Module $lockModule.Name -RequiredVersion $lockModule.keepVersion -AllowClobber -Scope CurrentUser
                    }

                    # Show latest version
                    Find-Module $lockModule.Name | Format-Table
                }

                $allVersions = Get-InstalledModule -Name $module.Name -AllVersions
    
                foreach ( $version in $allVersions ) {
                    write-host "Checking locked $($module.Name) $($version.Version)"
                    if ( $version.Version -ne $lockModule.keepVersion ) {
                        Write-Host "Removing $($module.Name) $($version.Version)" -ForegroundColor Green
                        Uninstall-Module -Name $module.Name -RequiredVersion $version.Version -Force
                    }
                }
            }
        }

        if ( $versionLocked -eq $False ) {
            Update-Module -Name $module.Name
        }
    }
    
    # Cleanup older versions of modules
    Write-Host "Cleanup older versions of modules" -ForegroundColor Cyan
    $modules = Get-InstalledModule
    foreach ($module in $modules) {
        Get-InstalledModule -Name $module.Name -AllVersions | Group-Object Name | ForEach-Object {
            $moduleGroup = $_.Group | Sort-Object { [version]$_.Version } -Descending
            $latest = $moduleGroup | Select-Object -First 1
            $olderVersions = $moduleGroup | Where-Object { $_.Version -ne $latest.Version }
            foreach ($old in $olderVersions) {
                Write-Host "Removing $($old.Name) v$($old.Version)" -ForegroundColor Green
                Uninstall-Module -Name $old.Name -RequiredVersion $old.Version -Force
            }
        }
    }

    # Check orphaned modules
    $installed = Get-InstalledModule
    $modulePaths = $env:PSModulePath -split ';'
    $modulePaths = $modulePaths | Where-Object {
        $_ -notmatch '\\WindowsPowerShell\\v1\.0\\Modules\\?$'
    }

    $basemodules = @(
        "CimCmdlets",
        "Microsoft.PowerShell.Archive",
        "Microsoft.PowerShell.Diagnostics",
        "Microsoft.PowerShell.Host",
        "Microsoft.PowerShell.Management",
        "Microsoft.PowerShell.PSResourceGet",
        "Microsoft.PowerShell.Security",
        "Microsoft.PowerShell.ThreadJob",
        "Microsoft.PowerShell.Utility",
        "Microsoft.WSMan.Management",
        "PackageManagement",
        "PowerShellGet",
        "PSDiagnostics",
        "PSReadLine",
        "Microsoft.PowerShell.Operation.Validation",
        "Pester"
    )

    foreach ($path in $modulePaths) {
        Get-ChildItem -Path $path -Directory -ErrorAction:SilentlyContinue | ForEach-Object {
            if ($installed.Name -notcontains $_.Name -and $basemodules -notcontains $_.Name) {
                Write-Host "Potential orphaned module: $($_.FullName)" -ForegroundColor Yellow
            }
        }
    }

    # Problems Sharepoint Online not updating
    $getmodule = Get-InstalledModule -Name Microsoft.Online.SharePoint.PowerShell -ErrorAction:SilentlyContinue
    if($getmodule)
    {
        $findmodule = Find-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction:SilentlyContinue
        if($getmodule -and $findmodule -and $getmodule.Version -ne $findmodule.Version)
        {
            # Uninstalling promatic module module
            Write-Host Removing Microsoft.Online.SharePoint.PowerShell to fix update issue. -ForegroundColor Green
            Uninstall-Module Microsoft.Online.SharePoint.PowerShell -Force -AllVersions -ErrorAction:SilentlyContinue
            $profilemodulepath = (Split-Path $PROFILE) + "\Modules"
            $spPath = Join-Path $profilemodulepath 'Microsoft.Online.SharePoint.PowerShell'
            if (Test-Path $spPath)
            {
                Remove-Item $spPath -Recurse -Force
            }
        }
    }
}

Function ConnectEXOnlineJSON {
    $JSON = Get-Content "$PSScriptRoot/Tenants.json" -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
    if (-not $JSON) {
        $JSON = [PSCustomObject]@{}
    }

    if (-not $JSON.Tenants) {
        $JSON | Add-Member -MemberType NoteProperty -Name Tenants -Value (@())
        #$JSON.Tenants = @()
    }
    $tenants = @($JSON.Tenants)

    if($tenants.Count -eq 0)
    {
        Write-Host "No Tenants defined. Run JSONentries" -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
    }

    for ($i = 0; $i -lt $tenants.Count; $i++) {
        Write-Host "$($i+1). $($tenants[$i].Name)"
    }

    $choice = Read-Host "Select tenant"
    if ($choice -notmatch '^\d+$' -or [int]$choice -lt 1 -or [int]$choice -gt $tenants.Count) {
        Write-Host "Invalid selection" -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
    }
    $selectedTenant = $tenants[$choice - 1]


    Write-Host Available services: 'MSGraph','MSGraphBeta','MSTeams','SharePointOnline','SharePointPnP','SecAndCompCenter','ExchangeOnline','MSEntra' -ForegroundColor Cyan
    $Service = (Read-Host -Prompt "Which service (leave empty for all)?").Trim()

    if($Service -eq ""){
        & $PSScriptRoot/ConnectO365Services.ps1 -TenantId $selectedTenant.TenantId -AppId $selectedTenant.AppId -CertificateThumbprint $selectedTenant.CertThumbprint
    }
    else{
        & $PSScriptRoot/ConnectO365Services.ps1 -TenantId $selectedTenant.TenantId -AppId $selectedTenant.AppId -CertificateThumbprint $selectedTenant.CertThumbprint -Services $Service
    }

    GetCurrentAccounts
}

Function ConnectEXOnlineMFA {
    Write-Host Available services: 'MSGraph','MSGraphBeta','MSTeams','SharePointOnline','SharePointPnP','SecAndCompCenter','ExchangeOnline','MSEntra' -ForegroundColor Cyan
    $Service = (Read-Host -Prompt "Which service (leave empty for all)?").Trim()

    if($Service -eq ""){
        & $PSScriptRoot/ConnectO365Services.ps1 -MFA -GraphScopes $Script:GraphScopes -EntraScopes $Script:EntraScopes
    }
    else{
        & $PSScriptRoot/ConnectO365Services.ps1 -MFA -GraphScopes $Script:GraphScopes -EntraScopes $Script:EntraScopes -Services $Service
    }

    GetCurrentAccounts
}

Function DisconnectEXOnline {
    & $PSScriptRoot/ConnectO365Services.ps1 -Disconnect
    $Script:ConnectedTenantName = ""

    Write-Host All disconnected
}

Function GetCurrentAccounts {
    $MgOrg = Get-MgOrganization -ErrorAction SilentlyContinue
    $Script:ConnectedTenantName = $MgOrg.DisplayName
    Write-Host MSGraph: $MgOrg.DisplayName

    try {
        $GetEntraUser = (Get-EntraUser -Top 1 -ErrorAction:Stop).UserPrincipalName
    }
    catch {
        $GetEntraUser = $null
    }
    write-host Entra: $GetEntraUser

    Write-Host Teams: (Get-CsTenant -ErrorAction:SilentlyContinue).DisplayName

    try {
        $GetSPOSite = (Get-SPOSite -Limit 1 -ErrorAction:Stop).Url
    }
    catch {
        $GetSPOSite = $null
    }
    Write-Host SharePointOnline: $GetSPOSite

    try {
        $GetPnPConnection = (Get-PnPConnection -ErrorAction:Stop).Url
    }
    catch {
        $GetPnPConnection = $null
    }
    Write-Host SharePointPnP: $GetPnPConnection

    Write-Host "SecAndCompCenter & ExchangeOnline (returns 2 values if both connected): " (Get-ConnectionInformation -ErrorAction:SilentlyContinue).UserPrincipalName
}

function JSONentries{
    $JSON = Get-Content "$PSScriptRoot/Tenants.json" -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
    if (-not $JSON) {
        $JSON = [PSCustomObject]@{}
    }

    if (-not $JSON.Tenants) {
        $JSON | Add-Member -MemberType NoteProperty -Name Tenants -Value (@())
        #$JSON.Tenants = @()
    }
    $JSON.Tenants = @($JSON.Tenants)

    # Haal huidige tenant op
    $Org = Get-MgOrganization -ErrorAction SilentlyContinue
    $CurrentTenantName = $Org.DisplayName

    # Zoek bestaande entry op tenantnaam
    $ExistingTenantIndex = -1

    for ($i = 0; $i -lt $JSON.Tenants.Count; $i++) {
        if ($JSON.Tenants[$i].TenantId -eq $Org.Id) {
            $ExistingTenantIndex = $i
            break
        }
    }

    Write-Host ""
    Write-Host "Current tenant: $CurrentTenantName" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. New entry"


    if ($ExistingTenantIndex -ge 0) {
        Write-Host "2. Overwrite existing entry"
    }
    else {
        Write-Host "2. Overwrite existing entry - unavailable, no matching tenant name found" -ForegroundColor DarkGray
    }

    Write-Host "3. Delete entry"
    Write-Host ""

    $choice = Read-Host "Select action"

    if ($choice -notmatch '^[1-3]$') {
        Write-Host "Invalid selection" -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
    }

    if ($choice -eq '2' -and $ExistingTenantIndex -lt 0) {
        Write-Host "Overwrite is not available because no existing JSON entry matches tenant name '$CurrentTenantName'." -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
    }

    if ($choice -eq '3') {
        if (-not $JSON.Tenants -or $JSON.Tenants.Count -eq 0) {
            Write-Host "No JSON entries found to delete." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return
        }

        Write-Host ""
        Write-Host "Select entry to delete:" -ForegroundColor Cyan

        for ($i = 0; $i -lt $JSON.Tenants.Count; $i++) {
            Write-Host "$($i + 1). $($JSON.Tenants[$i].Name) - $($JSON.Tenants[$i].TenantId)"
        }

        Write-Host "0. Cancel"
        Write-Host ""

        $deleteChoice = Read-Host "Select entry"

        if ($deleteChoice -eq '0') {
            Write-Host "Delete cancelled" -ForegroundColor Yellow
            return
        }

        if (
            $deleteChoice -notmatch '^\d+$' -or
            [int]$deleteChoice -lt 1 -or
            [int]$deleteChoice -gt $JSON.Tenants.Count
        ) {
            Write-Host "Invalid selection" -ForegroundColor Red
            Start-Sleep -Seconds 2
            return
        }

        $deleteIndex = [int]$deleteChoice - 1
        $deletedEntry = $JSON.Tenants[$deleteIndex]

        $JSON.Tenants = @(
            for ($i = 0; $i -lt $JSON.Tenants.Count; $i++) {
                if ($i -ne $deleteIndex) {
                    $JSON.Tenants[$i]
                }
            }
        )

        $JSON | ConvertTo-Json | Set-Content -Path "$PSScriptRoot/Tenants.json"

        Write-Host "Deleted JSON entry: $($deletedEntry.Name)" -ForegroundColor Yellow
        return
    }


    # Start create application

    # certificate
    $cert = Get-ChildItem -Path Cert:\CurrentUser\My\ | where { $_.subject -eq "CN=$($Shell.WindowTitle)" } -ErrorAction SilentlyContinue
    if (-not $cert)
    {
        $cert = New-SelfSignedCertificate `
            -Subject "CN=$($Shell.WindowTitle)" `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -KeySpec KeyExchange `
            -KeyLength 2048 `
            -NotAfter $Script:CertificateValidityPeriod

        Write-Host "Certificate generated" -ForegroundColor Yellow
    }
    else
    {
        Write-Host "Existing certificate found" -ForegroundColor Yellow
    }
    Export-Certificate -Cert $cert -FilePath "$PSScriptRoot\$($Shell.WindowTitle -replace '[\\/:*?""<>|]', '_').cer"

    $CertThumb = $cert.Thumbprint


    # Create the application with the necessary permissions
    $app = Get-MgApplication -ConsistencyLevel eventual -Filter "DisplayName eq '$($Shell.WindowTitle)'"
    if (-not $app) {
        $app = New-MgApplication -DisplayName "$($Shell.WindowTitle)"
        $app = HELPER_Wait-ForObject `
            -WaitingFor "application" `
            -ScriptBlock {
                Get-MgApplication `
                    -Filter "appId eq '$($app.AppId)'" `
                    -ErrorAction SilentlyContinue |
                    Select-Object -First 1
            }

        $appSp = New-MgServicePrincipal -AppId $app.AppId
        $appSp = HELPER_Wait-ForObject `
            -WaitingFor "client service principal" `
            -ScriptBlock {
                Get-MgServicePrincipal `
                    -Filter "appId eq '$($app.AppId)'" `
                    -ErrorAction SilentlyContinue |
                    Select-Object -First 1
            }

        Write-Host "New App created" -ForegroundColor Yellow

        Write-Host "Pause for Entra to propagate" -ForegroundColor Yellow
        Start-Sleep -Seconds 10
    }
    else
    {
        Write-Host "Existing App found" -ForegroundColor Yellow
    }

    # Check whether certificate already exists
    $existingCert = $app.KeyCredentials | Where-Object {
        $_.CustomKeyIdentifier -and
        ([System.Convert]::ToBase64String($_.CustomKeyIdentifier).ToUpperInvariant() -eq $CertThumb.ToUpperInvariant())
    }

    if ($existingCert) {
        Write-Host "Certificate already exists on app registration" -ForegroundColor Yellow
    }
    else {
        $newKeyCredential = @{
            Type          = "AsymmetricX509Cert"
            Usage         = "Verify"
            Key           = $cert.RawData
            DisplayName   = "$Env:ComputerName $([Environment]::UserName)"
            StartDateTime = $cert.NotBefore
            EndDateTime   = $cert.NotAfter
            CustomKeyIdentifier = [System.Convert]::FromBase64String($Certificate.Thumbprint)
        }

        $updatedKeyCredentials = @($app.KeyCredentials) + $newKeyCredential

        Update-MgApplication `
            -ApplicationId $app.Id `
            -KeyCredentials $updatedKeyCredentials
        Write-Host "Certificate imported to App" -ForegroundColor Yellow
    }


    $requiredResourceAccess = @()
    $graphResourceAccess = @()
    $exoResourceAccess = @()
    $spoResourceAccess = @()

    # Get existing service principals
    if (-not $appSp) {
        $appSp = Get-MgServicePrincipal `
            -Filter "appId eq '$($app.AppId)'"
    }

    $appSpAra = Get-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $appSp.Id

    # Graph service principal
    $graphSp = Get-MgServicePrincipal `
        -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

    $permissions = $Script:GraphScopes

    foreach ($permission in $permissions) {
        $appRole = $graphSp.AppRoles |
            Where-Object {
                $_.Value -eq $permission -and
                $_.AllowedMemberTypes -contains "Application"
            }

        if (-not $appRole) {
            Write-Host "Permission not found or not available as application permission: $permission" -ForegroundColor Red
            continue
        }

        $existingAssignment = $appSpAra |
            Where-Object {
                $_.ResourceId -eq $graphSp.Id -and
                $_.AppRoleId -eq $appRole.Id
            }

        $graphResourceAccess += @{
            Id   = $appRole.Id
            Type = "Role"
        }
 
        if (-not $existingAssignment) {
            $dummy = New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $appSp.Id `
                -PrincipalId $appSp.Id `
                -ResourceId $graphSp.Id `
                -AppRoleId $appRole.Id

            Write-Host "New Graph permission added: $permission" -ForegroundColor Yellow
        }
        else
        {
            Write-Host "Graph permission already exists: $permission" -ForegroundColor Yellow
        }
    }

    $requiredResourceAccess += @{
        ResourceAppId  = $graphSp.AppId
        ResourceAccess = @($graphResourceAccess)
    }


    # Exchange service principal
    $exoSp = Get-MgServicePrincipal `
        -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'"

    $permissions = @(
#        "full_access_as_app", # old EXO permission. Should not be needed
        "Exchange.AdminAPI.ManageAsApp",
        "Exchange.ManageAsApp",
        "Exchange.ManageAsAppV2"
    )

    foreach ($permission in $permissions) {
        $appRole = $exoSp.AppRoles |
            Where-Object {
                $_.Value -eq $permission -and
                $_.AllowedMemberTypes -contains "Application"
            }

        if (-not $appRole) {
            Write-Host "Permission not found or not available as application permission: $permission" -ForegroundColor Red
            continue
        }

        $existingAssignment = $appSpAra |
            Where-Object {
                $_.ResourceId -eq $exoSp.Id -and
                $_.AppRoleId -eq $appRole.Id
            }

        $exoResourceAccess += @{
            Id   = $appRole.Id
            Type = "Role"
        }
 
        if (-not $existingAssignment) {
            $dummy = New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $appSp.Id `
                -PrincipalId $appSp.Id `
                -ResourceId $exoSp.Id `
                -AppRoleId $appRole.Id

            Write-Host "New EXO permission added: $permission" -ForegroundColor Yellow
        }
        else
        {
            Write-Host "EXO permission already exists: $permission" -ForegroundColor Yellow
        }
    }

    $requiredResourceAccess += @{
        ResourceAppId  = $exoSp.AppId
        ResourceAccess = @($exoResourceAccess)
    }


    # Sharepoint service principal
    $spoSp = Get-MgServicePrincipal `
        -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'"

    $permissions = @(
        "Sites.FullControl.All",
#        "AllSites.FullControl", # Delegated permission not available for applications
        "TermStore.ReadWrite.All",
        "User.ReadWrite.All"
    )

    foreach ($permission in $permissions) {
        $appRole = $spoSp.AppRoles |
            Where-Object {
                $_.Value -eq $permission -and
                $_.AllowedMemberTypes -contains "Application"
            }

        if (-not $appRole) {
            Write-Host "Permission not found or not available as application permission: $permission" -ForegroundColor Red
            continue
        }

        $existingAssignment = $appSpAra |
            Where-Object {
                $_.ResourceId -eq $spoSp.Id -and
                $_.AppRoleId -eq $appRole.Id
            }

        $spoResourceAccess += @{
            Id   = $appRole.Id
            Type = "Role"
        }

        if (-not $existingAssignment) {
            $dummy = New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $appSp.Id `
                -PrincipalId $appSp.Id `
                -ResourceId $spoSp.Id `
                -AppRoleId $appRole.Id

            Write-Host "New Sharepoint permission added: $permission" -ForegroundColor Yellow
        }
        else
        {
            Write-Host "Sharepoint permission already exists: $permission" -ForegroundColor Yellow
        }
    }

    $requiredResourceAccess += @{
        ResourceAppId  = $spoSp.AppId
        ResourceAccess = @($spoResourceAccess)
    }


    # Write configured permissions to app registration
    Update-MgApplication `
        -ApplicationId $app.Id `
        -RequiredResourceAccess $requiredResourceAccess
    Write-Host "App updated with configured permissions" -ForegroundColor Yellow


    # Get Global Administrator role
    $role = Get-MgDirectoryRole |
        Where-Object { $_.DisplayName -eq "Global Administrator" }

    if (-not $role) {
        Write-Host "Global Administrator directory role was not found. Role may not be activated." -ForegroundColor Red
        return
    }
    else {
        # Check if service principal is already a member
        $existingMember = Get-MgDirectoryRoleMemberAsServicePrincipal -DirectoryRoleId $role.Id -All |
            Where-Object { $_.Id -eq $appSp.Id }

        if (-not $existingMember) {
            $confirmGA = Read-Host "Add service principal to Global Administrator? [Y/N]"

            if ($confirmGA -ne 'y') {
                Write-Host "Skipping Global Administrator assignment" -ForegroundColor Yellow
            }
            else {
                Write-Host "Adding service principal to Global Administrator role..." -ForegroundColor Yellow

                New-MgDirectoryRoleMemberByRef `
                    -DirectoryRoleId $role.Id `
                    -BodyParameter @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($appSp.Id)"
                    }
            }

        }
        else {
            Write-Host "Service principal is already a member of the Global Administrator role." -ForegroundColor Yellow
        }
    }

    # End create application

    $Entry = @{
        Name = $Org.DisplayName
        TenantId = $Org.Id
        AppId = $app.AppId
        CertThumbprint = $CertThumb
    }
    if ($choice -eq '1' -or $ExistingTenantIndex -lt 0)
    {
        $JSON.Tenants += $Entry
        Write-Host "JSON entry added" -ForegroundColor Yellow
    }
    else{
        $JSON.Tenants[$ExistingTenantIndex] = $Entry
        Write-Host "JSON entry modified: $($Org.DisplayName)" -ForegroundColor Yellow
    }

    $JSON | ConvertTo-Json | Set-Content -Path "$PSScriptRoot/Tenants.json"
}

function Reload-Profile {
    Write-Host "Reloading PowerShell profile from $PROFILE" -ForegroundColor Cyan
    . $PROFILE
}

function HELPER_Show-FunctionMenu {
    $filePath = $PSCommandPath
    $functionNames = @()

    if (Test-Path $filePath) {
        $content = Get-Content $filePath -Raw
        $matches = [regex]::Matches($content, '(?im)^\s*function\s+([a-zA-Z0-9_-]+)')

        foreach ($match in $matches) {
            $functionNames += $match.Groups[1].Value
        }
    }

    if (-not $functionNames) {
        Write-Host "No functions found in profile script." -ForegroundColor Red
        return
    }
    else
    {
        $functionNames = $functionNames | Where-Object { -not $_.StartsWith("HELPER_") }

        $functionNames += "Exit"
    }

    while ($true) {
        Clear-Host
        Write-Host "==== PowerShell Function Menu ====" -ForegroundColor Cyan
        if ($Script:ConnectedTenantName -ne "" -and $Script:ConnectedTenantName -ne $null) {
            Write-Host "Current Tenant: $($Script:ConnectedTenantName)`n" -ForegroundColor Green
        }
        Write-Host "Select a function to run:`n"

        for ($i = 0; $i -lt $functionNames.Count; $i++) {
            Write-Host "$($i + 1). $($functionNames[$i])"
        }

        $selection = Read-Host "`nEnter your choice (number)"
        if ($selection -match '^\d+$' -and [int]$selection -gt 0 -and [int]$selection -le $functionNames.Count) {
            $choice = $functionNames[$selection - 1]

            if ($choice -eq 'Exit') {
                break
            }

            #try {
                Write-Host "`nRunning function: $choice" -ForegroundColor Yellow
                & $choice
            #}
            #catch {
            #    Write-Error "Error executing function: $_"
            #}

            Write-Host "`nPress Enter to return to the menu..."
            [void][System.Console]::ReadLine()
        }
        else {
            Write-Host "Invalid selection. Try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }

        $functionNames = $functionNames
    }
}

function HELPER_Wait-ForObject {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [int]$TimeoutSeconds = 60,
        [int]$DelaySeconds = 3,
        [string]$WaitingFor = "object"
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    do {
        $result = & $ScriptBlock

        if ($result) {
            return $result
        }

        Start-Sleep -Seconds $DelaySeconds
    }
    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds)

    throw "Timed out waiting for $WaitingFor."
}

function HELPER_Test-NoPowerShellArguments {
    $cmd = [Environment]::CommandLine.Trim()

    if ([string]::IsNullOrWhiteSpace($cmd)) {
        return $false
    }

    # Remove first token: quoted executable path or unquoted executable path
    if ($cmd -match '^\s*"[^"]+"\s*(?<args>.*)$') {
        $remaining = $Matches.args.Trim()
    }
    elseif ($cmd -match '^\s*\S+\s*(?<args>.*)$') {
        $remaining = $Matches.args.Trim()
    }
    else {
        return $false
    }

    if ($remaining -eq "-WorkingDirectory ~") {
        return $true
    }

    return [string]::IsNullOrWhiteSpace($remaining)
}

if (HELPER_Test-NoPowerShellArguments) {
    HELPER_Show-FunctionMenu
}
