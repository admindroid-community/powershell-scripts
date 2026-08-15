function Get-M365Task {
    <#
        .SYNOPSIS
        Lists or searches registered toolkit tasks.

        .EXAMPLE
        Get-M365Task -Category Security

        .EXAMPLE
        Get-M365Task -Search 'forward'

        .EXAMPLE
        Get-M365Task -Id 'licensing.unused-licenses'
    #>
    [CmdletBinding(DefaultParameterSetName = 'Filter')]
    param(
        [Parameter(ParameterSetName = 'ById', Position = 0)]
        [string]$Id,

        [Parameter(ParameterSetName = 'Filter')]
        [string]$Category,

        [Parameter(ParameterSetName = 'Filter')]
        [string]$Search,

        [Parameter(ParameterSetName = 'Filter')]
        [ValidateSet('Graph', 'ExchangeOnline', 'PnP', 'Teams')]
        [string]$Service,

        [Parameter(ParameterSetName = 'Filter')]
        [ValidateSet('ReadOnly', 'Write', 'Destructive')]
        [string]$Risk
    )

    $all = $script:TaskRegistry.Values

    if ($PSCmdlet.ParameterSetName -eq 'ById') {
        if (-not $script:TaskRegistry.Contains($Id)) {
            Write-Error "No task registered with Id '$Id'. Run Get-M365Task to list available tasks."
            return
        }
        return $script:TaskRegistry[$Id]
    }

    $results = $all
    if ($Category) { $results = $results | Where-Object { $_.Category -like "*$Category*" } }
    if ($Service)  { $results = $results | Where-Object { $Service -in $_.Service } }
    if ($Risk)     { $results = $results | Where-Object { $_.Risk -eq $Risk } }
    if ($Search) {
        $results = $results | Where-Object {
            $_.Name -like "*$Search*" -or
            $_.Synopsis -like "*$Search*" -or
            $_.Id -like "*$Search*" -or
            ($_.Replaces -and $_.Replaces -like "*$Search*")
        }
    }

    $results | Sort-Object Category, Name
}
