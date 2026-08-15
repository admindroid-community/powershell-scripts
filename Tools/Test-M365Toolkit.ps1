<#
.SYNOPSIS
    Smoke tests for the M365Toolkit module. No tenant connection required.

.DESCRIPTION
    Validates the parts that can be checked without live Microsoft 365 credentials:
    module import, task registration and validation, parameter coercion, the risk gate,
    the execution envelope, and export. Connectivity is stubbed inside the module scope.

.EXAMPLE
    pwsh -File ./Tools/Test-M365Toolkit.ps1
#>
[CmdletBinding()]
param(
    [string]$ModulePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'M365Toolkit/M365Toolkit.psd1')
)

$script:Pass = 0
$script:Fail = 0

function Test-Case {
    param([string]$Name, [scriptblock]$Body)
    try {
        & $Body
        $script:Pass++
        Write-Host "  PASS  $Name" -ForegroundColor Green
    } catch {
        $script:Fail++
        Write-Host "  FAIL  $Name" -ForegroundColor Red
        Write-Host "        $($_.Exception.Message)" -ForegroundColor DarkGray
    }
}

function Assert-True {
    param($Condition, [string]$Message = 'Expected condition to be true')
    if (-not $Condition) { throw $Message }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    if ($Expected -ne $Actual) {
        if ($Message) { throw $Message }
        throw "Expected '$Expected' but got '$Actual'"
    }
}

Write-Host "`nM365Toolkit smoke tests" -ForegroundColor Cyan
Write-Host "Module: $ModulePath`n"

Import-Module $ModulePath -Force -ErrorAction Stop
$mod = Get-Module M365Toolkit

# ---------------------------------------------------------------- registry
Write-Host "Registry" -ForegroundColor Yellow

Test-Case 'module exports the documented commands' {
    $cmds = (Get-Command -Module M365Toolkit).Name
    foreach ($c in 'Connect-M365Toolkit','Invoke-M365Task','Get-M365Task','Register-M365Task','Export-M365Result','Show-M365Console') {
        Assert-True ($c -in $cmds) "Missing exported command: $c"
    }
}

Test-Case 'tasks are registered at import' {
    Assert-True ((Get-M365Task).Count -gt 0) 'No tasks registered'
}

Test-Case 'every task has a unique dotted id' {
    $ids = (Get-M365Task).Id
    Assert-Equal $ids.Count (@($ids | Select-Object -Unique).Count) 'Duplicate task ids found'
    foreach ($id in $ids) {
        Assert-True ($id -match '^[a-z0-9]+(\.[a-z0-9-]+)+$') "Malformed id: $id"
    }
}

Test-Case 'every task declares a valid service and risk level' {
    foreach ($t in (Get-M365Task)) {
        foreach ($s in @($t.Service)) {
            Assert-True ($s -in @('Graph','ExchangeOnline','PnP','Teams')) "Task $($t.Id) has bad service '$s'"
        }
        Assert-True ($t.Risk -in @('ReadOnly','Write','Destructive')) "Task $($t.Id) has bad risk '$($t.Risk)'"
    }
}

Test-Case 'every task Execute is a scriptblock' {
    foreach ($t in (Get-M365Task)) {
        Assert-True ($t.Execute -is [scriptblock]) "Task $($t.Id) Execute is not a scriptblock"
    }
}

Test-Case 'write and destructive tasks expose a WhatIf parameter' {
    foreach ($t in (Get-M365Task | Where-Object { $_.Risk -ne 'ReadOnly' })) {
        $names = @($t.Parameters | ForEach-Object { $_.Name })
        Assert-True ('WhatIf' -in $names) "Risky task $($t.Id) has no WhatIf parameter"
    }
}

Test-Case 'choice parameters declare options' {
    foreach ($t in (Get-M365Task)) {
        foreach ($p in @($t.Parameters)) {
            if ($p.Type -in @('Choice','MultiChoice')) {
                Assert-True (@($p.Options).Count -gt 0) "Task $($t.Id) parameter $($p.Name) has no Options"
            }
        }
    }
}

Test-Case 'choice defaults are valid options' {
    foreach ($t in (Get-M365Task)) {
        foreach ($p in @($t.Parameters)) {
            if ($p.Type -eq 'Choice' -and $p.ContainsKey('Default') -and $null -ne $p.Default) {
                Assert-True ($p.Default -in $p.Options) "Task $($t.Id) parameter $($p.Name) default '$($p.Default)' is not in Options"
            }
        }
    }
}

Test-Case 'registering a malformed task is rejected' {
    $threw = $false
    try { Register-M365Task @{ Id = 'BAD ID'; Name = 'x'; Category = 'c'; Synopsis = 's'; Service = 'Graph'; Execute = {} } }
    catch { $threw = $true }
    Assert-True $threw 'Malformed task was accepted'
}

Test-Case 'duplicate registration without -Force is rejected' {
    $threw = $false
    try { Register-M365Task @{ Id = 'identity.guest-users'; Name = 'dupe'; Category = 'c'; Synopsis = 's'; Service = 'Graph'; Execute = {} } }
    catch { $threw = $true }
    Assert-True $threw 'Duplicate task id was accepted'
}

# ------------------------------------------------------------- parameters
Write-Host "`nParameter binding" -ForegroundColor Yellow

Test-Case 'defaults are applied when nothing is supplied' {
    $t = Get-M365Task -Id 'licensing.unused-licenses'
    $r = & $mod { param($t) Resolve-M365TaskParameter -Task $t -Supplied @{} } $t
    Assert-Equal 90 $r.InactiveDays
}

Test-Case 'supplied values override defaults and are coerced to type' {
    $t = Get-M365Task -Id 'licensing.unused-licenses'
    $r = & $mod { param($t) Resolve-M365TaskParameter -Task $t -Supplied @{ InactiveDays = '45' } } $t
    Assert-Equal 45 $r.InactiveDays
    Assert-True ($r.InactiveDays -is [int]) 'InactiveDays was not coerced to int'
}

Test-Case 'missing required parameter throws' {
    $t = Get-M365Task -Id 'identity.user-membership'
    $threw = $false
    try { & $mod { param($t) Resolve-M365TaskParameter -Task $t -Supplied @{} } $t } catch { $threw = $true }
    Assert-True $threw 'Missing required parameter was accepted'
}

Test-Case 'invalid choice value throws' {
    $t = Get-M365Task -Id 'security.risky-users'
    $threw = $false
    try { & $mod { param($t) Resolve-M365TaskParameter -Task $t -Supplied @{ MinimumRiskLevel = 'nope' } } $t } catch { $threw = $true }
    Assert-True $threw 'Invalid choice value was accepted'
}

# --------------------------------------------------------------- executor
Write-Host "`nExecutor" -ForegroundColor Yellow

Test-Case 'running a task while disconnected throws' {
    $threw = $false
    try { Invoke-M365Task -Id 'identity.guest-users' -ErrorAction Stop | Out-Null } catch { $threw = $true }
    Assert-True $threw 'Disconnected run did not throw'
}

# Stub connectivity inside the module so the executor can be exercised offline.
# These must be written into the module's *script* scope: a plain `function` inside
# & $mod { } lands in a child scope that is discarded when the scriptblock returns.
& $mod {
    Set-Item -Path function:script:Test-M365Connection   -Value { param($Service) return $true }
    Set-Item -Path function:script:Assert-M365Connection -Value { param($Service, $Scopes) return }
    Set-Item -Path function:script:Test-M365Scope        -Value { param($RequiredScopes) return @() }
}

Register-M365Task -Force @{
    Id       = 'test.echo'
    Name     = 'Test Echo'
    Category = 'Test'
    Synopsis = 'Emits a fixed number of rows for testing.'
    Service  = 'Graph'
    Risk     = 'ReadOnly'
    Parameters = @(
        @{ Name = 'RowCount'; Type = 'Int'; Default = 3; Help = 'Rows to emit.' }
    )
    Execute = {
        param($P)
        1..([int]$P.RowCount) | ForEach-Object {
            [pscustomobject]@{ Index = $_; Label = "row-$_"; Squared = $_ * $_ }
        }
    }
}

Test-Case 'task emits objects to the pipeline' {
    $rows = @(Invoke-M365Task -Id 'test.echo' -Parameters @{ RowCount = 5 })
    Assert-Equal 5 $rows.Count
    Assert-Equal 'row-3' $rows[2].Label
    Assert-Equal 9 $rows[2].Squared
}

Test-Case 'results are composable with standard cmdlets' {
    $big = @(Invoke-M365Task -Id 'test.echo' -Parameters @{ RowCount = 10 } | Where-Object Squared -gt 50)
    Assert-Equal 3 $big.Count
}

Test-Case '-Detailed returns an execution envelope' {
    $env = Invoke-M365Task -Id 'test.echo' -Parameters @{ RowCount = 4 } -Detailed
    Assert-Equal 'test.echo' $env.TaskId
    Assert-Equal 4 $env.RowCount
    Assert-True $env.Succeeded 'Envelope reports failure'
    Assert-True ($env.Duration.TotalMilliseconds -ge 0) 'No duration recorded'
}

Test-Case 'a failing task is captured in the envelope rather than thrown' {
    Register-M365Task -Force @{
        Id = 'test.boom'; Name = 'Test Boom'; Category = 'Test'
        Synopsis = 'Always throws.'; Service = 'Graph'; Risk = 'ReadOnly'
        Execute = { param($P) throw 'intentional failure' }
    }
    $env = Invoke-M365Task -Id 'test.boom' -Detailed
    Assert-True (-not $env.Succeeded) 'Envelope did not report failure'
    Assert-True ($env.Error -match 'intentional failure') "Unexpected error text: $($env.Error)"
}

Test-Case 'unknown task id throws' {
    $threw = $false
    try { Invoke-M365Task -Id 'nope.missing' -ErrorAction Stop | Out-Null } catch { $threw = $true }
    Assert-True $threw 'Unknown task id was accepted'
}

Test-Case 'risky task is skipped when confirmation is declined' {
    Register-M365Task -Force @{
        Id = 'test.risky'; Name = 'Test Risky'; Category = 'Test'
        Synopsis = 'Write operation.'; Service = 'Graph'; Risk = 'Write'
        Parameters = @(@{ Name = 'WhatIf'; Type = 'Bool'; Default = $true; Help = 'x' })
        Execute = { param($P) [pscustomobject]@{ Ran = $true } }
    }
    # -WhatIf makes ShouldProcess return false, so the task body must not run.
    $out = @(Invoke-M365Task -Id 'test.risky' -WhatIf)
    Assert-Equal 0 $out.Count 'Risky task ran despite -WhatIf'
}

Test-Case 'risky task runs when forced' {
    $out = @(Invoke-M365Task -Id 'test.risky' -Force)
    Assert-Equal 1 $out.Count
    Assert-True $out[0].Ran 'Task body did not execute'
}

# ----------------------------------------------------------------- export
Write-Host "`nExport" -ForegroundColor Yellow

Test-Case 'CSV export round-trips' {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) "m365tk-$(Get-Random).csv"
    try {
        Invoke-M365Task -Id 'test.echo' -Parameters @{ RowCount = 6 } | Export-M365Result -Path $tmp
        Assert-True (Test-Path $tmp) 'CSV was not created'
        $back = Import-Csv $tmp
        Assert-Equal 6 @($back).Count
        Assert-Equal 'row-1' $back[0].Label
    } finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
}

Test-Case 'HTML export is produced' {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) "m365tk-$(Get-Random).html"
    try {
        Invoke-M365Task -Id 'test.echo' -Parameters @{ RowCount = 2 } |
            Export-M365Result -Path $tmp -Format Html -Title 'Test Report'
        Assert-True (Test-Path $tmp) 'HTML was not created'
        $content = Get-Content $tmp -Raw
        Assert-True ($content -match 'Test Report') 'Title missing from HTML'
        Assert-True ($content -match 'row-1') 'Data missing from HTML'
    } finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
}

# ------------------------------------------------------------------ logging
Write-Host "`nLogging" -ForegroundColor Yellow

Test-Case 'task execution is logged' {
    Clear-M365Log -Confirm:$false
    Invoke-M365Task -Id 'test.echo' -Parameters @{ RowCount = 1 } | Out-Null
    $log = Get-M365Log
    Assert-True (@($log).Count -gt 0) 'Nothing was logged'
    Assert-True (@($log | Where-Object { $_.Source -eq 'Executor' }).Count -gt 0) 'No executor entries logged'
}

# ------------------------------------------------------------------ result
Write-Host ""
Write-Host ("{0} passed, {1} failed" -f $script:Pass, $script:Fail) -ForegroundColor $(if ($script:Fail) { 'Red' } else { 'Green' })
Write-Host ""

if ($script:Fail -gt 0) { exit 1 }
