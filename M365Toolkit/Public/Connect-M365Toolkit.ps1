function Connect-M365Toolkit {
    <#
        .SYNOPSIS
        Connects to one or more Microsoft 365 services using a single shared context.

        .DESCRIPTION
        Certificate-based auth is the supported path for unattended/scheduled runs.
        Omit -CertificateThumbprint for an interactive sign-in.

        The username/password path used by the original script collection is not
        implemented: Exchange Online no longer accepts basic authentication.

        .PARAMETER Service
        Which services to connect. Defaults to Graph.

        .PARAMETER TenantId
        Directory (tenant) ID. Required for certificate auth.

        .PARAMETER ClientId
        Application (client) ID of the app registration. Required for certificate auth.

        .PARAMETER CertificateThumbprint
        Thumbprint of a certificate in CurrentUser\My or LocalMachine\My. Triggers
        unattended auth.

        .PARAMETER Organization
        Tenant domain (contoso.onmicrosoft.com). Required for Exchange Online and PnP
        certificate auth.

        .PARAMETER SharePointUrl
        SharePoint admin or site URL. Required when connecting PnP.

        .PARAMETER Scopes
        Delegated Graph scopes to request during interactive sign-in.

        .PARAMETER InstallMissingModules
        Install any required module that is not present, scoped to the current user.

        .EXAMPLE
        Connect-M365Toolkit -Service Graph, ExchangeOnline `
            -TenantId $tid -ClientId $cid -CertificateThumbprint $thumb `
            -Organization contoso.onmicrosoft.com

        .EXAMPLE
        Connect-M365Toolkit -Service Graph -Scopes 'User.Read.All','AuditLog.Read.All'
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Graph', 'ExchangeOnline', 'PnP', 'Teams')]
        [string[]]$Service = @('Graph'),

        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [string]$Organization,
        [string]$SharePointUrl,
        [string[]]$Scopes,

        [switch]$InstallMissingModules,

        [string]$LogFile
    )

    if ($LogFile) {
        $script:LogFilePath = $LogFile
        Write-M365Log -Level Info -Source 'Connection' -Message "Logging to $LogFile"
    }

    $useCert = -not [string]::IsNullOrWhiteSpace($CertificateThumbprint)
    if ($useCert -and (-not $TenantId -or -not $ClientId)) {
        throw "Certificate authentication requires both -TenantId and -ClientId."
    }

    $context = @{
        TenantId              = $TenantId
        ClientId              = $ClientId
        CertificateThumbprint = $CertificateThumbprint
        Organization          = $Organization
        SharePointUrl         = $SharePointUrl
    }

    foreach ($svc in $Service) {
        if (-not (Test-M365Module -Service $svc -InstallMissing:$InstallMissingModules)) {
            $moduleName = $script:ServiceModules[$svc].Name
            throw "Cannot connect to $svc - module '$moduleName' is not installed. Re-run with -InstallMissingModules, or: Install-Module $moduleName -Scope CurrentUser"
        }

        try {
            Connect-M365Service -Service $svc -Context $context -Scopes $Scopes
        } catch {
            Write-M365Log -Level Error -Source 'Connection' -Message "$svc connection failed: $($_.Exception.Message)"
            throw
        }
    }

    # Stored so tasks can silently reconnect an expired session.
    $script:AuthContext = $context

    Get-M365Connection
}
