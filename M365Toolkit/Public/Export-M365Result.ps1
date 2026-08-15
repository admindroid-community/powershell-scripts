function Export-M365Result {
    <#
        .SYNOPSIS
        Exports task results to CSV, JSON or a self-contained HTML report.

        .DESCRIPTION
        Export is a caller decision here, not something buried inside each task. That
        is the main structural difference from the original scripts, which wrote CSV
        from inside the script body.

        .EXAMPLE
        Invoke-M365Task -Id 'exchange.mailbox-sizes' | Export-M365Result -Path .\sizes.xlsx.csv

        .EXAMPLE
        Invoke-M365Task -Id 'security.risky-users' |
            Export-M365Result -Path .\risky.html -Format Html -Title 'Risky Users'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object[]]$InputObject,

        [Parameter(Mandatory, Position = 0)]
        [string]$Path,

        [ValidateSet('Csv', 'Json', 'Html', 'Auto')]
        [string]$Format = 'Auto',

        [string]$Title = 'Microsoft 365 Report',

        [switch]$PassThru
    )

    begin {
        $collected = [System.Collections.Generic.List[object]]::new()
    }

    process {
        foreach ($item in $InputObject) {
            if ($null -ne $item) { $collected.Add($item) }
        }
    }

    end {
        if ($collected.Count -eq 0) {
            Write-Warning "Nothing to export - the task returned no rows."
            return
        }

        if ($Format -eq 'Auto') {
            $Format = switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
                '.json' { 'Json' }
                '.html' { 'Html' }
                '.htm'  { 'Html' }
                default { 'Csv' }
            }
        }

        $dir = Split-Path -Parent $Path
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        if (-not $PSCmdlet.ShouldProcess($Path, "Export $($collected.Count) row(s) as $Format")) {
            return
        }

        switch ($Format) {
            'Csv'  { $collected | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 }
            'Json' { $collected | ConvertTo-Json -Depth 6 | Set-Content -Path $Path -Encoding UTF8 }
            'Html' {
                $style = @'
<style>
  body { font-family: Segoe UI, system-ui, sans-serif; margin: 2rem; color: #1a1a1a; background: #fff; }
  h1 { font-size: 1.4rem; margin-bottom: .25rem; }
  .meta { color: #666; font-size: .85rem; margin-bottom: 1.5rem; }
  table { border-collapse: collapse; width: 100%; font-size: .85rem; }
  th { background: #f3f4f6; text-align: left; padding: .5rem .6rem; border-bottom: 2px solid #d1d5db; position: sticky; top: 0; }
  td { padding: .4rem .6rem; border-bottom: 1px solid #eee; }
  tr:nth-child(even) td { background: #fafafa; }
</style>
'@
                $meta = "<div class='meta'>$($collected.Count) row(s) &middot; generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>"
                $collected |
                    ConvertTo-Html -Title $Title -PreContent "<h1>$Title</h1>$meta" -Head $style |
                    Set-Content -Path $Path -Encoding UTF8
            }
        }

        Write-M365Log -Level Success -Source 'Export' -Message "Wrote $($collected.Count) row(s) to $Path"

        if ($PassThru) { $collected.ToArray() }
    }
}
