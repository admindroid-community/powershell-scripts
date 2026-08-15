<#
.SYNOPSIS
    Launches the Microsoft 365 Engineering Console.

.DESCRIPTION
    Convenience launcher: imports the bundled M365Toolkit module from this repository
    and opens the GUI. Requires Windows (WPF).

    For command-line use, import the module directly:
        Import-Module ./M365Toolkit/M365Toolkit.psd1
        Get-M365Task
        Invoke-M365Task -Id 'licensing.unused-licenses'

.EXAMPLE
    ./Start-M365Console.ps1

.EXAMPLE
    # Create a desktop shortcut target:
    powershell.exe -ExecutionPolicy Bypass -File "C:\path\to\Start-M365Console.ps1"
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$manifest = Join-Path $PSScriptRoot 'M365Toolkit/M365Toolkit.psd1'
if (-not (Test-Path $manifest)) {
    throw "M365Toolkit not found at '$manifest'. Run this script from inside the repository."
}

Write-Host "Loading M365Toolkit..." -ForegroundColor Cyan
Import-Module $manifest -Force

$taskCount = (Get-M365Task).Count
Write-Host "$taskCount task(s) available." -ForegroundColor Green

Show-M365Console
