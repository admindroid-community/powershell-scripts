@{
    RootModule        = 'M365Toolkit.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '6f2b1c84-9d3e-4a17-9c5d-2b7e8f1a4d60'
    Author            = 'M365 Engineering'
    Description       = 'Task-driven Microsoft 365 administration toolkit with a WPF console. Modern Graph / Exchange Online cmdlets, certificate-based auth, pipeline-first output.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        'Connect-M365Toolkit'
        'Disconnect-M365Toolkit'
        'Get-M365Connection'
        'Get-M365Task'
        'Invoke-M365Task'
        'Register-M365Task'
        'Export-M365Result'
        'Show-M365Console'
        # Exported because the GUI's worker runspace calls them at top-level scope.
        'Get-M365Log'
        'Clear-M365Log'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData = @{
        PSData = @{
            Tags       = @('Microsoft365','Graph','ExchangeOnline','Entra','Reporting','GUI')
            ProjectUri = 'https://github.com/miggy7474/powershell-scripts-forked'
        }
    }
}
