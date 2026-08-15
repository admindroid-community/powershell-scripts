<#
    Small shared helpers used across task definitions.
#>

function Get-SafeProperty {
    <#
        .SYNOPSIS
        Reads a property that may not exist on the object, returning a default instead
        of throwing.

        .DESCRIPTION
        Graph and Exchange objects vary by tenant licensing and API version. This keeps
        a missing optional property from aborting a 40,000-object scan.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$Name,

        $Default = $null
    )

    if ($null -eq $InputObject) { return $Default }

    # Support dotted paths such as 'SignInActivity.LastSignInDateTime'
    $current = $InputObject
    foreach ($segment in $Name.Split('.')) {
        if ($null -eq $current) { return $Default }

        if ($current -is [hashtable] -or $current -is [System.Collections.IDictionary]) {
            if ($current.Contains($segment)) { $current = $current[$segment] } else { return $Default }
            continue
        }

        $prop = $current.PSObject.Properties[$segment]
        if ($null -eq $prop) { return $Default }
        $current = $prop.Value
    }

    if ($null -eq $current) { return $Default }
    return $current
}

function ConvertTo-FriendlySize {
    <#
        .SYNOPSIS
        Formats a byte count as GB with two decimals.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Bytes)

    if ($null -eq $Bytes) { return $null }
    try {
        return [math]::Round(([double]$Bytes) / 1GB, 2)
    } catch {
        return $null
    }
}

function ConvertFrom-ExchangeSize {
    <#
        .SYNOPSIS
        Parses an Exchange size string such as '1.234 GB (1,325,400,064 bytes)' into bytes.

        .DESCRIPTION
        Exchange Online returns sizes as display strings when the session is serialized
        over PowerShell remoting, so the numeric value has to be recovered from the text.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Size)

    if ($null -eq $Size) { return $null }

    $text = $Size.ToString()
    if ($text -match '\(([\d,]+)\s*bytes\)') {
        return [int64]($Matches[1] -replace ',', '')
    }

    # Some sessions return a real ByteQuantifiedSize object.
    $bytes = Get-SafeProperty -InputObject $Size -Name 'Value' -Default $null
    if ($null -ne $bytes) {
        try { return [int64]$bytes.ToBytes() } catch { }
    }

    return $null
}

function Get-DaysSince {
    <#
        .SYNOPSIS
        Whole days between a timestamp and now; $null when the timestamp is absent.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Timestamp)

    if ($null -eq $Timestamp -or $Timestamp -eq '') { return $null }
    try {
        $dt = if ($Timestamp -is [datetime]) { $Timestamp } else { [datetime]::Parse($Timestamp) }
        return [int]((Get-Date) - $dt).TotalDays
    } catch {
        return $null
    }
}

function Test-ExternalDomain {
    <#
        .SYNOPSIS
        True when an SMTP address sits outside the tenant's accepted domains.
    #>
    [CmdletBinding()]
    param(
        [string]$Address,
        [string[]]$AcceptedDomains
    )

    if ([string]::IsNullOrWhiteSpace($Address)) { return $false }
    if ($Address -notmatch '@') { return $false }

    $domain = ($Address -split '@')[-1].Trim().TrimEnd('>').ToLowerInvariant()
    return ($domain -notin ($AcceptedDomains | ForEach-Object { $_.ToLowerInvariant() }))
}

function Get-M365AcceptedDomain {
    <#
        .SYNOPSIS
        Cached list of the tenant's accepted domains, for external-recipient checks.
    #>
    [CmdletBinding()]
    param([switch]$Refresh)

    if ($script:AcceptedDomainCache -and -not $Refresh) {
        return $script:AcceptedDomainCache
    }

    $domains = @()
    if (Get-Command Get-AcceptedDomain -ErrorAction SilentlyContinue) {
        $domains = Get-AcceptedDomain -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DomainName
    } elseif (Get-Command Get-MgDomain -ErrorAction SilentlyContinue) {
        $domains = Get-MgDomain -All -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
    }

    $script:AcceptedDomainCache = @($domains)
    return $script:AcceptedDomainCache
}
