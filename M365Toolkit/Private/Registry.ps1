<#
    Task definition validation and parameter binding.

    A task is a hashtable with a fixed shape. Validating at registration time means a
    malformed task fails loudly at import rather than halfway through a tenant scan.
#>

$script:ValidRiskLevels = @('ReadOnly', 'Write', 'Destructive')
$script:ValidServices   = @('Graph', 'ExchangeOnline', 'PnP', 'Teams')
$script:ValidParamTypes = @('String', 'Int', 'Bool', 'DateTime', 'Choice', 'MultiChoice', 'File')

function Test-M365TaskDefinition {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Definition)

    $errors = [System.Collections.Generic.List[string]]::new()

    foreach ($key in 'Id', 'Name', 'Category', 'Synopsis', 'Service', 'Execute') {
        if (-not $Definition.ContainsKey($key) -or $null -eq $Definition[$key]) {
            $errors.Add("Missing required key '$key'.")
        }
    }
    if ($errors.Count -gt 0) {
        return $errors
    }

    if ($Definition.Id -notmatch '^[a-z0-9]+(\.[a-z0-9-]+)+$') {
        $errors.Add("Id '$($Definition.Id)' must be dotted lowercase, e.g. 'licensing.unused-licenses'.")
    }

    foreach ($svc in @($Definition.Service)) {
        if ($svc -notin $script:ValidServices) {
            $errors.Add("Service '$svc' is not one of: $($script:ValidServices -join ', ').")
        }
    }

    $risk = if ($Definition.ContainsKey('Risk')) { $Definition.Risk } else { 'ReadOnly' }
    if ($risk -notin $script:ValidRiskLevels) {
        $errors.Add("Risk '$risk' is not one of: $($script:ValidRiskLevels -join ', ').")
    }

    if ($Definition.Execute -isnot [scriptblock]) {
        $errors.Add("Execute must be a scriptblock.")
    }

    if ($Definition.ContainsKey('Parameters') -and $Definition.Parameters) {
        foreach ($p in $Definition.Parameters) {
            if ($p -isnot [hashtable]) { $errors.Add("Each entry in Parameters must be a hashtable."); continue }
            if (-not $p.Name) { $errors.Add("A parameter is missing 'Name'."); continue }

            $ptype = if ($p.ContainsKey('Type')) { $p.Type } else { 'String' }
            if ($ptype -notin $script:ValidParamTypes) {
                $errors.Add("Parameter '$($p.Name)' has invalid Type '$ptype'. Valid: $($script:ValidParamTypes -join ', ').")
            }
            if ($ptype -in @('Choice', 'MultiChoice') -and -not $p.Options) {
                $errors.Add("Parameter '$($p.Name)' is type $ptype but declares no Options.")
            }
        }
    }

    return $errors
}

function Resolve-M365TaskParameter {
    <#
        .SYNOPSIS
        Coerces caller-supplied values to the types a task declares, applying defaults
        and enforcing Required.
    #>
    [CmdletBinding()]
    param(
        # A registered task object (pscustomobject), not the raw definition hashtable.
        [Parameter(Mandatory)]$Task,
        [hashtable]$Supplied = @{}
    )

    $resolved = @{}
    $declared = @(Get-SafeProperty -InputObject $Task -Name 'Parameters' -Default @())

    foreach ($p in $declared) {
        $name  = $p.Name
        $ptype = if ($p.ContainsKey('Type')) { $p.Type } else { 'String' }
        $hasValue = $Supplied.ContainsKey($name) -and
                    $null -ne $Supplied[$name] -and
                    -not ($Supplied[$name] -is [string] -and [string]::IsNullOrWhiteSpace($Supplied[$name]))

        if (-not $hasValue) {
            if ($p.ContainsKey('Default') -and $null -ne $p.Default) {
                $resolved[$name] = $p.Default
                continue
            }
            if ($p.Required) {
                throw "Task '$($Task.Id)': required parameter '$name' was not supplied."
            }
            $resolved[$name] = $null
            continue
        }

        $raw = $Supplied[$name]

        try {
            switch ($ptype) {
                'Int'      { $resolved[$name] = [int]$raw }
                'Bool'     { $resolved[$name] = [bool]$raw }
                'DateTime' { $resolved[$name] = if ($raw -is [datetime]) { $raw } else { [datetime]::Parse($raw) } }
                'MultiChoice' {
                    $vals = @($raw)
                    foreach ($v in $vals) {
                        if ($v -notin $p.Options) { throw "'$v' is not a valid option for '$name'." }
                    }
                    $resolved[$name] = $vals
                }
                'Choice' {
                    if ($raw -notin $p.Options) { throw "'$raw' is not a valid option for '$name'." }
                    $resolved[$name] = $raw
                }
                default    { $resolved[$name] = [string]$raw }
            }
        } catch {
            throw "Task '$($Task.Id)': parameter '$name' - $($_.Exception.Message)"
        }
    }

    # Pass through anything extra the caller supplied but the task did not declare.
    foreach ($k in $Supplied.Keys) {
        if (-not $resolved.ContainsKey($k)) { $resolved[$k] = $Supplied[$k] }
    }

    return $resolved
}
