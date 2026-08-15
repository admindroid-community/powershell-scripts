<#
    Single shared authentication layer.

    Replaces the ~150 copy-pasted Connect_MgGraph / Connect_EXO helpers in the original
    script collection. Certificate-based auth is the primary path; interactive is the
    fallback for ad-hoc console use. The retired basic-auth (username/password) path is
    deliberately NOT implemented - Exchange Online no longer accepts it.
#>

# Module requirements per service.
$script:ServiceModules = @{
    Graph          = @{ Name = 'Microsoft.Graph.Authentication'; Extra = @('Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement') }
    ExchangeOnline = @{ Name = 'ExchangeOnlineManagement';       Extra = @() }
    PnP            = @{ Name = 'PnP.PowerShell';                 Extra = @() }
    Teams          = @{ Name = 'MicrosoftTeams';                 Extra = @() }
}

function Test-M365Module {
    <#
        .SYNOPSIS
        Verifies the module backing a service is available, without prompting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Graph', 'ExchangeOnline', 'PnP', 'Teams')]
        [string]$Service,

        [switch]$InstallMissing
    )

    $spec = $script:ServiceModules[$Service]
    $name = $spec.Name

    if (Get-Module -Name $name -ListAvailable -ErrorAction SilentlyContinue) {
        return $true
    }

    if (-not $InstallMissing) {
        Write-M365Log -Level Error -Source 'Connection' -Message "Module '$name' is not installed (required for $Service)."
        return $false
    }

    Write-M365Log -Level Info -Source 'Connection' -Message "Installing module '$name' for current user..."
    try {
        Install-Module -Name $name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-M365Log -Level Success -Source 'Connection' -Message "Installed '$name'."
        return $true
    } catch {
        Write-M365Log -Level Error -Source 'Connection' -Message "Failed to install '$name': $($_.Exception.Message)"
        return $false
    }
}

function Connect-M365Service {
    <#
        .SYNOPSIS
        Connects one service using the stored auth context.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Graph', 'ExchangeOnline', 'PnP', 'Teams')]
        [string]$Service,

        [Parameter(Mandatory)]
        [hashtable]$Context,

        [string[]]$Scopes
    )

    $useCert = -not [string]::IsNullOrWhiteSpace($Context.CertificateThumbprint)

    switch ($Service) {

        'Graph' {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            if ($useCert) {
                Connect-MgGraph -TenantId $Context.TenantId `
                                -ClientId $Context.ClientId `
                                -CertificateThumbprint $Context.CertificateThumbprint `
                                -NoWelcome -ErrorAction Stop
            } else {
                $connectScopes = if ($Scopes) { $Scopes } else { @('User.Read.All', 'Directory.Read.All') }
                Connect-MgGraph -Scopes $connectScopes -NoWelcome -ErrorAction Stop
            }
            $ctx = Get-MgContext
            $script:ConnectionState.Graph = [pscustomobject]@{
                Connected   = $true
                Account     = if ($ctx) { $ctx.Account } else { $null }
                TenantId    = if ($ctx) { $ctx.TenantId } else { $Context.TenantId }
                AuthType    = if ($useCert) { 'Certificate' } else { 'Delegated' }
                Scopes      = if ($ctx) { $ctx.Scopes } else { @() }
                ConnectedAt = Get-Date
            }
        }

        'ExchangeOnline' {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            if ($useCert) {
                if ([string]::IsNullOrWhiteSpace($Context.Organization)) {
                    throw "Exchange Online certificate auth requires -Organization (e.g. contoso.onmicrosoft.com)."
                }
                Connect-ExchangeOnline -AppId $Context.ClientId `
                                       -CertificateThumbprint $Context.CertificateThumbprint `
                                       -Organization $Context.Organization `
                                       -ShowBanner:$false -ErrorAction Stop
            } else {
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            }
            $script:ConnectionState.ExchangeOnline = [pscustomobject]@{
                Connected    = $true
                Organization = $Context.Organization
                AuthType     = if ($useCert) { 'Certificate' } else { 'Delegated' }
                ConnectedAt  = Get-Date
            }
        }

        'PnP' {
            Import-Module PnP.PowerShell -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($Context.SharePointUrl)) {
                throw "PnP requires -SharePointUrl (e.g. https://contoso-admin.sharepoint.com)."
            }
            if ($useCert) {
                Connect-PnPOnline -Url $Context.SharePointUrl `
                                  -ClientId $Context.ClientId `
                                  -Thumbprint $Context.CertificateThumbprint `
                                  -Tenant $Context.Organization -ErrorAction Stop
            } else {
                Connect-PnPOnline -Url $Context.SharePointUrl -Interactive -ErrorAction Stop
            }
            $script:ConnectionState.PnP = [pscustomobject]@{
                Connected   = $true
                Url         = $Context.SharePointUrl
                AuthType    = if ($useCert) { 'Certificate' } else { 'Delegated' }
                ConnectedAt = Get-Date
            }
        }

        'Teams' {
            Import-Module MicrosoftTeams -ErrorAction Stop
            if ($useCert) {
                Connect-MicrosoftTeams -TenantId $Context.TenantId `
                                       -ApplicationId $Context.ClientId `
                                       -CertificateThumbprint $Context.CertificateThumbprint `
                                       -ErrorAction Stop | Out-Null
            } else {
                Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
            }
            $script:ConnectionState.Teams = [pscustomobject]@{
                Connected   = $true
                AuthType    = if ($useCert) { 'Certificate' } else { 'Delegated' }
                ConnectedAt = Get-Date
            }
        }
    }

    Write-M365Log -Level Success -Source 'Connection' -Message "Connected to $Service."
}

function Test-M365Connection {
    <#
        .SYNOPSIS
        Returns $true if the named service currently looks connected.

        .DESCRIPTION
        Probes the live session rather than trusting cached state, because tokens
        expire and other code may have disconnected underneath us.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Graph', 'ExchangeOnline', 'PnP', 'Teams')]
        [string]$Service
    )

    try {
        switch ($Service) {
            'Graph' {
                if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) { return $false }
                return $null -ne (Get-MgContext -ErrorAction SilentlyContinue)
            }
            'ExchangeOnline' {
                if (-not (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) { return $false }
                $info = Get-ConnectionInformation -ErrorAction SilentlyContinue
                return [bool]($info | Where-Object { $_.State -eq 'Connected' })
            }
            'PnP' {
                if (-not (Get-Command Get-PnPConnection -ErrorAction SilentlyContinue)) { return $false }
                return $null -ne (Get-PnPConnection -ErrorAction SilentlyContinue)
            }
            'Teams' {
                if (-not (Get-Command Get-CsTenant -ErrorAction SilentlyContinue)) { return $false }
                return $null -ne (Get-CsTenant -ErrorAction SilentlyContinue)
            }
        }
    } catch {
        return $false
    }
    return $false
}

function Assert-M365Connection {
    <#
        .SYNOPSIS
        Ensures the services a task needs are connected, reconnecting from the stored
        auth context where possible.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Service,

        [string[]]$Scopes
    )

    foreach ($svc in $Service) {
        if (Test-M365Connection -Service $svc) { continue }

        if (-not $script:AuthContext) {
            throw "Not connected to $svc. Run Connect-M365Toolkit first."
        }

        Write-M365Log -Level Info -Source 'Connection' -Message "$svc session not active - reconnecting."
        Connect-M365Service -Service $svc -Context $script:AuthContext -Scopes $Scopes
    }
}

function Test-M365Scope {
    <#
        .SYNOPSIS
        Warns when the current Graph token is missing scopes a task declares.

        .DESCRIPTION
        Non-fatal by design: app-only tokens carry roles rather than scopes, and the
        task itself will surface a clearer error if it genuinely lacks permission.
    #>
    [CmdletBinding()]
    param([string[]]$RequiredScopes)

    if (-not $RequiredScopes -or $RequiredScopes.Count -eq 0) { return @() }
    if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) { return @() }

    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx -or -not $ctx.Scopes) { return @() }

    $missing = $RequiredScopes | Where-Object { $_ -notin $ctx.Scopes }
    if ($missing) {
        Write-M365Log -Level Warning -Source 'Connection' -Message "Token may be missing scope(s): $($missing -join ', ')"
    }
    return $missing
}
