function Disconnect-M365Toolkit {
    <#
        .SYNOPSIS
        Disconnects from Microsoft 365 services and clears the stored auth context.

        .EXAMPLE
        Disconnect-M365Toolkit
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [ValidateSet('Graph', 'ExchangeOnline', 'PnP', 'Teams')]
        [string[]]$Service = @('Graph', 'ExchangeOnline', 'PnP', 'Teams')
    )

    foreach ($svc in $Service) {
        if (-not (Test-M365Connection -Service $svc)) { continue }
        if (-not $PSCmdlet.ShouldProcess($svc, 'Disconnect')) { continue }

        try {
            switch ($svc) {
                'Graph'          { Disconnect-MgGraph -ErrorAction Stop | Out-Null }
                'ExchangeOnline' { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop }
                'PnP'            { Disconnect-PnPOnline -ErrorAction Stop }
                'Teams'          { Disconnect-MicrosoftTeams -ErrorAction Stop }
            }
            $script:ConnectionState[$svc] = $null
            Write-M365Log -Level Success -Source 'Connection' -Message "Disconnected from $svc."
        } catch {
            Write-M365Log -Level Warning -Source 'Connection' -Message "Disconnect from $svc reported: $($_.Exception.Message)"
        }
    }

    $script:AuthContext = $null
}
