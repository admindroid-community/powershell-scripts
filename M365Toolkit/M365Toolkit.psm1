<#
    M365Toolkit - module loader.

    Dot-sources Private, then Public, then every task definition under Tasks/.
    Task files call Register-M365Task, so the registry is populated at import time.
#>

# Strict mode is deliberately NOT enabled: Graph and Exchange return objects whose
# property sets vary by tenant licensing, and strict mode turns a missing optional
# property into a hard failure mid-scan. Tasks use Get-SafeProperty instead.

$script:ModuleRoot = $PSScriptRoot

# ---------------------------------------------------------------------------
# Module-scope state
# ---------------------------------------------------------------------------

# Registry of all known tasks, keyed by task Id.
$script:TaskRegistry = [ordered]@{}

# Tracks which services we believe we are currently connected to.
$script:ConnectionState = @{
    Graph          = $null
    ExchangeOnline = $null
    PnP            = $null
    Teams          = $null
}

# Populated by Connect-M365Toolkit so tasks can reconnect silently if needed.
$script:AuthContext = $null

# In-memory log ring buffer, surfaced in the GUI.
$script:LogBuffer = [System.Collections.Generic.List[object]]::new()

# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

foreach ($scope in 'Private', 'Public') {
    $dir = Join-Path $script:ModuleRoot $scope
    if (Test-Path $dir) {
        Get-ChildItem -Path $dir -Filter *.ps1 -File | Sort-Object Name | ForEach-Object {
            try {
                . $_.FullName
            } catch {
                throw "Failed to load $scope file '$($_.Name)': $_"
            }
        }
    }
}

# Task definitions last - they depend on Register-M365Task being defined.
$taskDir = Join-Path $script:ModuleRoot 'Tasks'
if (Test-Path $taskDir) {
    Get-ChildItem -Path $taskDir -Filter *.tasks.ps1 -File | Sort-Object Name | ForEach-Object {
        try {
            . $_.FullName
        } catch {
            Write-Warning "Failed to load task file '$($_.Name)': $_"
        }
    }
}

Write-Verbose "M365Toolkit loaded with $($script:TaskRegistry.Count) task(s)."
