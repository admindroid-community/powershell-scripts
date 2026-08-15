function Invoke-M365Task {
    <#
        .SYNOPSIS
        Runs a registered task and emits its results as objects.

        .DESCRIPTION
        Handles everything the original scripts each re-implemented: connection checks,
        reconnection, scope validation, confirmation gating for write operations, timing
        and error capture.

        Output goes to the pipeline as objects - not Write-Host, not straight to CSV -
        so results can be filtered, joined or exported by the caller.

        .PARAMETER Id
        Task id, e.g. 'licensing.unused-licenses'.

        .PARAMETER Parameters
        Hashtable of task parameters.

        .PARAMETER Detailed
        Return a result envelope (results plus timing, errors, row count) instead of
        the bare result objects. Used by the GUI.

        .PARAMETER Force
        Skip the confirmation prompt for Write and Destructive tasks.

        .EXAMPLE
        Invoke-M365Task -Id 'licensing.unused-licenses' -Parameters @{ InactiveDays = 90 }

        .EXAMPLE
        Invoke-M365Task -Id 'exchange.mailbox-sizes' | Where-Object TotalSizeGB -gt 40

        .EXAMPLE
        Invoke-M365Task -Id 'security.block-forwarding' -Parameters @{ WhatIf = $true }
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipelineByPropertyName)]
        [Alias('TaskId')]
        [string]$Id,

        [Parameter(Position = 1)]
        [hashtable]$Parameters = @{},

        [switch]$Detailed,

        [switch]$Force
    )

    process {
        if (-not $script:TaskRegistry.Contains($Id)) {
            throw "No task registered with Id '$Id'. Run Get-M365Task to list available tasks."
        }
        $task = $script:TaskRegistry[$Id]

        # --- Risk gate -----------------------------------------------------
        if ($task.Risk -ne 'ReadOnly' -and -not $Force) {
            $action = "Run $($task.Risk) task"
            $target = "$($task.Name) [$($task.Id)]"
            if (-not $PSCmdlet.ShouldProcess($target, $action)) {
                Write-M365Log -Level Warning -Source 'Executor' -Message "Task '$Id' cancelled by user."
                return
            }
        }

        # --- Parameters ----------------------------------------------------
        $resolved = Resolve-M365TaskParameter -Task $task -Supplied $Parameters

        # --- Connections ---------------------------------------------------
        Assert-M365Connection -Service $task.Service -Scopes $task.Scopes
        $missingScopes = Test-M365Scope -RequiredScopes $task.Scopes

        # --- Execute -------------------------------------------------------
        Write-M365Log -Level Info -Source 'Executor' -Message "Running '$($task.Name)' [$Id]."
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        $results = [System.Collections.Generic.List[object]]::new()
        $failure = $null

        try {
            # The scriptblock receives the resolved parameter hashtable as $args[0].
            $output = & $task.Execute $resolved
            foreach ($o in $output) {
                if ($null -ne $o) { $results.Add($o) }
            }
            $sw.Stop()
            Write-M365Log -Level Success -Source 'Executor' `
                -Message "'$($task.Name)' returned $($results.Count) row(s) in $([math]::Round($sw.Elapsed.TotalSeconds,1))s."
        } catch {
            $sw.Stop()
            $failure = $_
            Write-M365Log -Level Error -Source 'Executor' -Message "'$($task.Name)' failed: $($_.Exception.Message)"
        }

        # --- Emit ----------------------------------------------------------
        if ($Detailed) {
            [pscustomobject]@{
                TaskId        = $task.Id
                TaskName      = $task.Name
                Category      = $task.Category
                Risk          = $task.Risk
                Parameters    = $resolved
                Results       = $results.ToArray()
                RowCount      = $results.Count
                Duration      = $sw.Elapsed
                Succeeded     = ($null -eq $failure)
                Error         = if ($failure) { $failure.Exception.Message } else { $null }
                MissingScopes = $missingScopes
                RanAt         = Get-Date
            }
        } else {
            if ($failure) { throw $failure }
            $results.ToArray()
        }
    }
}
