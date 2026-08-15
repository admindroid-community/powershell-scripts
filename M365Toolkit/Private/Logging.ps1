<#
    Structured logging. Everything the toolkit does lands in a ring buffer that the
    GUI reads, and optionally in a transcript file on disk.
#>

function Write-M365Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Debug')]
        [string]$Level = 'Info',

        [string]$Source = 'Toolkit'
    )

    $entry = [pscustomobject]@{
        Timestamp = Get-Date
        Level     = $Level
        Source    = $Source
        Message   = $Message
    }

    $script:LogBuffer.Add($entry)

    # Keep the buffer bounded so a long GUI session does not grow without limit.
    if ($script:LogBuffer.Count -gt 5000) {
        $script:LogBuffer.RemoveRange(0, 1000)
    }

    switch ($Level) {
        'Error'   { Write-Verbose "[ERROR]   $Source :: $Message" }
        'Warning' { Write-Verbose "[WARN]    $Source :: $Message" }
        'Success' { Write-Verbose "[OK]      $Source :: $Message" }
        'Debug'   { Write-Debug   "[DEBUG]   $Source :: $Message" }
        default   { Write-Verbose "[INFO]    $Source :: $Message" }
    }

    if ($script:LogFilePath) {
        $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2} :: {3}' -f $entry.Timestamp, $Level.ToUpper(), $Source, $Message
        try {
            Add-Content -Path $script:LogFilePath -Value $line -Encoding UTF8 -ErrorAction Stop
        } catch {
            # Never let logging failures break a running task.
        }
    }
}

function Get-M365Log {
    [CmdletBinding()]
    param(
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Debug')]
        [string[]]$Level,

        [int]$Last = 200
    )

    $items = $script:LogBuffer.ToArray()
    if ($Level) {
        $items = $items | Where-Object { $_.Level -in $Level }
    }
    $items | Select-Object -Last $Last
}

function Clear-M365Log {
    [CmdletBinding(SupportsShouldProcess)]
    param()
    if ($PSCmdlet.ShouldProcess('M365Toolkit log buffer', 'Clear')) {
        $script:LogBuffer.Clear()
    }
}
