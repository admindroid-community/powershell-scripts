<#
.SYNOPSIS
    Static audit of the PowerShell scripts in this repository.

.DESCRIPTION
    AST-based analysis (not grep) that reports:
      - parse failures
      - retired module / cmdlet usage (AzureAD, MSOnline)
      - retired basic-authentication code paths
      - hardcoded secrets and plaintext credential handling
      - interactive prompts that block scheduled execution
      - Write-Host-only output that cannot be consumed by a caller
      - O(n^2) array append inside loops
      - missing error handling

    This produced the findings in REVIEW.md.

.PARAMETER Path
    Repository root to scan. Defaults to the parent of this script's folder.

.PARAMETER Format
    Text (default), Json, or Csv.

.PARAMETER MinimumSeverity
    Only report at or above this severity.

.EXAMPLE
    ./Tools/Invoke-RepoAudit.ps1

.EXAMPLE
    ./Tools/Invoke-RepoAudit.ps1 -Format Json | Out-File audit.json

.EXAMPLE
    ./Tools/Invoke-RepoAudit.ps1 -MinimumSeverity Error
#>
[CmdletBinding()]
param(
    [string]$Path,
    [ValidateSet('Text','Json','Csv')][string]$Format = 'Text',
    [ValidateSet('Info','Warning','Error')][string]$MinimumSeverity = 'Info'
)

if (-not $Path) { $Path = Split-Path -Parent $PSScriptRoot }
if (-not (Test-Path $Path)) { throw "Path not found: $Path" }

$severityRank = @{ Info = 0; Warning = 1; Error = 2 }
$floor = $severityRank[$MinimumSeverity]

$retiredCmdlets = @(
    'Connect-AzureAD','Connect-MsolService','Get-AzureADUser','Get-MsolUser','Set-MsolUser',
    'Get-AzureADGroup','Get-AzureADGroupMember','Get-MsolAccountSku','Get-MsolGroup',
    'Set-MsolUserLicense','Get-AzureADSubscribedSku','Get-AzureADDevice','New-MsolUser',
    'Get-AzureADServicePrincipal','Get-AzureADApplication','Get-MsolDomain',
    'Get-MsolCompanyInformation','Get-AzureADDirectoryRole','Set-AzureADUser'
)
$retiredModules = @('AzureAD','AzureADPreview','MSOnline')

$findings = [System.Collections.Generic.List[object]]::new()

function Add-Finding {
    param($File, $Severity, $Rule, $Line, $Message)
    $findings.Add([pscustomobject]@{
        File = $File; Severity = $Severity; Rule = $Rule; Line = $Line; Message = $Message
    })
}

$files = Get-ChildItem -Path $Path -Recurse -Filter *.ps1 -File |
         Where-Object { $_.FullName -notmatch '[\\/]\.git[\\/]' }

foreach ($file in $files) {
    $rel = $file.FullName.Substring($Path.Length).TrimStart('/','\')
    $raw = Get-Content -LiteralPath $file.FullName -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrEmpty($raw)) { continue }

    $tokens = $null; $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($raw, [ref]$tokens, [ref]$parseErrors)

    foreach ($pe in @($parseErrors)) {
        Add-Finding $rel 'Error' 'ParseError' $pe.Extent.StartLineNumber $pe.Message
    }
    if (-not $ast) { continue }

    # --- retired cmdlets ---
    $commands = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
    $usedRetired = @($commands |
        ForEach-Object { $_.GetCommandName() } |
        Where-Object { $_ -and $_ -in $retiredCmdlets } |
        Select-Object -Unique)

    if ($usedRetired.Count -gt 0) {
        Add-Finding $rel 'Error' 'RetiredCmdlet' 0 "Uses retired cmdlets: $($usedRetired -join ', ')"
    }

    foreach ($m in $retiredModules) {
        if ($raw -match "(?im)(Import-Module|Install-Module)\s+[`"']?$([regex]::Escape($m))[`"']?([^\w.-]|$)") {
            Add-Finding $rel 'Error' 'RetiredModule' 0 "References retired module '$m'."
        }
    }

    # --- retired basic auth path ---
    if ($raw -match 'ConvertTo-SecureString.*-AsPlainText' -and $raw -match 'Connect-ExchangeOnline') {
        Add-Finding $rel 'Error' 'BasicAuthPath' 0 'Username/password path for Exchange Online; basic authentication is disabled.'
    } elseif ($raw -match 'ConvertTo-SecureString.*-AsPlainText') {
        Add-Finding $rel 'Warning' 'PlainTextCredential' 0 'Credential converted from plaintext.'
    }

    # --- hardcoded secrets ---
    if ($raw -match '(?im)^\s*\$(Password|Pwd|ClientSecret|AppSecret)\s*=\s*[''"][^''"$]{6,}[''"]') {
        Add-Finding $rel 'Error' 'HardcodedSecret' 0 'Possible hardcoded credential assigned as a literal string.'
    }

    # --- interactive prompts ---
    $readHost = @($commands | Where-Object { $_.GetCommandName() -eq 'Read-Host' }).Count
    if ($readHost -gt 0) {
        Add-Finding $rel 'Info' 'InteractivePrompt' 0 "Read-Host used $readHost time(s); blocks unattended execution."
    }

    # --- Write-Host heavy ---
    $writeHost = @($commands | Where-Object { $_.GetCommandName() -eq 'Write-Host' }).Count
    if ($writeHost -ge 15) {
        Add-Finding $rel 'Info' 'WriteHostHeavy' 0 "Write-Host used $writeHost time(s); output is not pipeline-consumable."
    }

    # --- array append in loop ---
    $appends = $ast.FindAll({ param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Operator -eq 'PlusEquals'
    }, $true)

    $inLoop = 0
    foreach ($a in $appends) {
        $p = $a.Parent
        while ($p) {
            if ($p -is [System.Management.Automation.Language.ForEachStatementAst] -or
                $p -is [System.Management.Automation.Language.ForStatementAst] -or
                $p -is [System.Management.Automation.Language.WhileStatementAst]) { $inLoop++; break }
            $p = $p.Parent
        }
    }
    if ($inLoop -gt 0) {
        Add-Finding $rel 'Info' 'ArrayAppendInLoop' 0 "$inLoop array '+=' append(s) inside a loop; O(n^2) growth."
    }

    # --- error handling ---
    $tryCount = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] }, $true)).Count
    if ($tryCount -eq 0 -and $raw -notmatch '-ErrorAction' -and $raw -notmatch '\$ErrorActionPreference' -and $raw.Length -gt 800) {
        Add-Finding $rel 'Warning' 'NoErrorHandling' 0 'No try/catch and no ErrorAction handling.'
    }
}

$filtered = $findings | Where-Object { $severityRank[$_.Severity] -ge $floor }

switch ($Format) {
    'Json' { $filtered | ConvertTo-Json -Depth 4 }
    'Csv'  { $filtered | ConvertTo-Csv -NoTypeInformation }
    'Text' {
        Write-Host ""
        Write-Host "Repository audit: $Path" -ForegroundColor Cyan
        Write-Host ("Scanned {0} script(s); {1} finding(s)" -f @($files).Count, @($filtered).Count)
        Write-Host ""

        foreach ($group in ($filtered | Group-Object Severity | Sort-Object { $severityRank[$_.Name] } -Descending)) {
            $colour = switch ($group.Name) { 'Error' { 'Red' } 'Warning' { 'Yellow' } default { 'Gray' } }
            Write-Host "$($group.Name.ToUpper()) ($($group.Count))" -ForegroundColor $colour

            foreach ($rule in ($group.Group | Group-Object Rule | Sort-Object Count -Descending)) {
                Write-Host ("  {0,-22} {1}" -f $rule.Name, $rule.Count)
                foreach ($f in ($rule.Group | Select-Object -First 5)) {
                    $loc = if ($f.Line -gt 0) { ":$($f.Line)" } else { '' }
                    Write-Host ("      {0}{1}" -f $f.File, $loc) -ForegroundColor DarkGray
                }
                if ($rule.Count -gt 5) { Write-Host ("      ... and $($rule.Count - 5) more") -ForegroundColor DarkGray }
            }
            Write-Host ""
        }
    }
}
