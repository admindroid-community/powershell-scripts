function Get-M365Connection {
    <#
        .SYNOPSIS
        Reports live connection status for each supported service.

        .EXAMPLE
        Get-M365Connection | Format-Table
    #>
    [CmdletBinding()]
    param()

    foreach ($svc in 'Graph', 'ExchangeOnline', 'PnP', 'Teams') {
        $live  = Test-M365Connection -Service $svc
        $state = $script:ConnectionState[$svc]

        [pscustomobject]@{
            Service     = $svc
            Connected   = $live
            AuthType    = if ($live -and $state) { $state.AuthType } else { $null }
            Identity    = if ($live -and $state) {
                              if ($state.PSObject.Properties.Name -contains 'Account')      { $state.Account }
                              elseif ($state.PSObject.Properties.Name -contains 'Organization') { $state.Organization }
                              elseif ($state.PSObject.Properties.Name -contains 'Url')      { $state.Url }
                              else { $null }
                          } else { $null }
            ConnectedAt = if ($live -and $state) { $state.ConnectedAt } else { $null }
        }
    }
}
