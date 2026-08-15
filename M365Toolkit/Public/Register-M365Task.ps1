function Register-M365Task {
    <#
        .SYNOPSIS
        Registers a task definition into the toolkit registry.

        .DESCRIPTION
        Called by every file under Tasks/. Adding a new capability to the toolkit and
        to the GUI means writing one of these - no GUI code changes required.

        A definition is a hashtable:

            Id         - dotted lowercase unique id, e.g. 'licensing.unused-licenses'
            Name       - display name
            Category   - GUI grouping
            Synopsis   - one-line description
            Service    - one or more of Graph, ExchangeOnline, PnP, Teams
            Scopes     - Graph permission scopes the task needs
            Risk       - ReadOnly (default) | Write | Destructive
            Replaces   - original repo folder this supersedes (optional)
            Parameters - array of parameter hashtables (optional)
            Execute    - scriptblock taking one hashtable argument, emitting objects

        Each parameter hashtable:

            Name     - parameter name
            Type     - String | Int | Bool | DateTime | Choice | MultiChoice | File
            Help     - shown in the GUI
            Default  - default value (optional)
            Required - $true to enforce (optional)
            Options  - required for Choice / MultiChoice

        .EXAMPLE
        Register-M365Task @{
            Id       = 'identity.guest-users'
            Name     = 'Guest User Report'
            Category = 'Identity'
            Synopsis = 'Lists guest accounts with sign-in activity.'
            Service  = 'Graph'
            Scopes   = @('User.Read.All')
            Execute  = { param($P) Get-MgUser -Filter "userType eq 'Guest'" -All }
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [hashtable]$Definition,

        [switch]$Force
    )

    process {
        $problems = Test-M365TaskDefinition -Definition $Definition
        if ($problems.Count -gt 0) {
            $id = if ($Definition.ContainsKey('Id')) { $Definition.Id } else { '<no id>' }
            throw "Invalid task definition '$id':`n  - $($problems -join "`n  - ")"
        }

        if ($script:TaskRegistry.Contains($Definition.Id) -and -not $Force) {
            throw "Task '$($Definition.Id)' is already registered. Use -Force to replace it."
        }

        # Normalise optional keys so consumers never have to test for them.
        $normalised = @{
            Id         = $Definition.Id
            Name       = $Definition.Name
            Category   = $Definition.Category
            Synopsis   = $Definition.Synopsis
            Service    = @($Definition.Service)
            Scopes     = if ($Definition.ContainsKey('Scopes'))     { @($Definition.Scopes) }     else { @() }
            Risk       = if ($Definition.ContainsKey('Risk'))       { $Definition.Risk }          else { 'ReadOnly' }
            Replaces   = if ($Definition.ContainsKey('Replaces'))   { $Definition.Replaces }      else { $null }
            Notes      = if ($Definition.ContainsKey('Notes'))      { $Definition.Notes }         else { $null }
            Parameters = if ($Definition.ContainsKey('Parameters')) { @($Definition.Parameters) } else { @() }
            Execute    = $Definition.Execute
        }

        # Stored as an object, not a hashtable: Select-Object, Format-Table and WPF
        # data binding all read properties, and none of them read hashtable keys.
        $script:TaskRegistry[$Definition.Id] = [pscustomobject]$normalised
        Write-M365Log -Level Debug -Source 'Registry' -Message "Registered task '$($Definition.Id)'."
    }
}
