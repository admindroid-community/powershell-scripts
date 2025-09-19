# Microsoft 365 Modules Verification and Testing Script
# This script verifies installation and tests basic functionality

param(
    [Parameter(Mandatory = $false)]
    [switch]$TestConnectivity,
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateReport,
    
    [Parameter(Mandatory = $false)]
    [string]$ReportPath = "c:\temp\M365ModulesReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
)

Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "Microsoft 365 Modules Verification Script" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan

# Define modules to verify
$ModulesToVerify = @(
    @{
        Name = 'Microsoft.Graph'
        TestCommands = @('Connect-MgGraph', 'Get-MgContext', 'Get-MgUser')
        Category = 'Core'
        Description = 'Microsoft Graph SDK'
    },
    @{
        Name = 'PnP.PowerShell'
        TestCommands = @('Connect-PnPOnline', 'Get-PnPContext')
        Category = 'SharePoint'
        Description = 'SharePoint PnP PowerShell'
    },
    @{
        Name = 'Microsoft.Online.SharePoint.PowerShell'
        TestCommands = @('Connect-SPOService', 'Get-SPOSite')
        Category = 'SharePoint'
        Description = 'SharePoint Online Management Shell'
    },
    @{
        Name = 'MicrosoftTeams'
        TestCommands = @('Connect-MicrosoftTeams', 'Get-Team')
        Category = 'Teams'
        Description = 'Microsoft Teams PowerShell'
    },
    @{
        Name = 'ExchangeOnlineManagement'
        TestCommands = @('Connect-ExchangeOnline', 'Get-ConnectionInformation')
        Category = 'Exchange'
        Description = 'Exchange Online Management'
    },
    @{
        Name = 'Microsoft.PowerApps.Administration.PowerShell'
        TestCommands = @('Add-PowerAppsAccount', 'Get-PowerAppEnvironment')
        Category = 'PowerPlatform'
        Description = 'PowerApps Administration'
    },
    @{
        Name = 'Microsoft.Graph.Intune'
        TestCommands = @('Connect-MSGraph', 'Get-IntuneManagedDevice')
        Category = 'Intune'
        Description = 'Microsoft Intune PowerShell'
    }
)

$results = @()

Write-Host "`nVerifying installed modules..." -ForegroundColor Yellow

foreach ($moduleInfo in $ModulesToVerify) {
    $moduleName = $moduleInfo.Name
    Write-Host "`nChecking $moduleName..." -ForegroundColor Cyan
    
    $result = [PSCustomObject]@{
        ModuleName = $moduleName
        Category = $moduleInfo.Category
        Description = $moduleInfo.Description
        Installed = $false
        Version = 'Not Installed'
        Location = 'N/A'
        CommandsAvailable = @()
        CommandsMissing = @()
        Status = 'Not Installed'
        ImportSuccess = $false
        LastUpdated = 'Unknown'
    }
    
    try {
        # Check if module is installed
        $installedModule = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue | 
                          Sort-Object Version -Descending | 
                          Select-Object -First 1
        
        if ($installedModule) {
            $result.Installed = $true
            $result.Version = $installedModule.Version.ToString()
            $result.Location = $installedModule.ModuleBase
            $result.Status = 'Installed'
            
            # Get last write time as update indicator
            try {
                $moduleManifest = Join-Path $installedModule.ModuleBase "$moduleName.psd1"
                if (Test-Path $moduleManifest) {
                    $result.LastUpdated = (Get-Item $moduleManifest).LastWriteTime.ToString('yyyy-MM-dd')
                }
            }
            catch {
                $result.LastUpdated = 'Unknown'
            }
            
            Write-Host "  ✓ Installed: v$($result.Version)" -ForegroundColor Green
            
            # Test module import
            try {
                Import-Module $moduleName -Force -ErrorAction Stop
                $result.ImportSuccess = $true
                Write-Host "  ✓ Import successful" -ForegroundColor Green
                
                # Test commands
                foreach ($command in $moduleInfo.TestCommands) {
                    if (Get-Command $command -ErrorAction SilentlyContinue) {
                        $result.CommandsAvailable += $command
                        Write-Host "  ✓ Command available: $command" -ForegroundColor Green
                    }
                    else {
                        $result.CommandsMissing += $command
                        Write-Host "  ✗ Command missing: $command" -ForegroundColor Red
                    }
                }
                
                if ($result.CommandsMissing.Count -eq 0) {
                    $result.Status = 'Fully Functional'
                }
                else {
                    $result.Status = 'Partially Functional'
                }
            }
            catch {
                $result.ImportSuccess = $false
                $result.Status = 'Import Failed'
                $result.CommandsMissing = $moduleInfo.TestCommands
                Write-Host "  ✗ Import failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Host "  ✗ Not installed" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "  ✗ Error checking module: $($_.Exception.Message)" -ForegroundColor Red
        $result.Status = 'Error'
    }
    
    $results += $result
}

# Display summary
Write-Host "`n===========================================" -ForegroundColor Cyan
Write-Host "VERIFICATION SUMMARY" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan

$installed = ($results | Where-Object { $_.Installed }).Count
$fullyFunctional = ($results | Where-Object { $_.Status -eq 'Fully Functional' }).Count
$total = $results.Count

Write-Host "Total modules checked: $total" -ForegroundColor White
Write-Host "Installed modules: $installed" -ForegroundColor Green
Write-Host "Fully functional: $fullyFunctional" -ForegroundColor Green

# Group by category
$categories = $results | Group-Object Category | Sort-Object Name

foreach ($category in $categories) {
    Write-Host "`n--- $($category.Name.ToUpper()) MODULES ---" -ForegroundColor Yellow
    
    foreach ($module in $category.Group) {
        $statusColor = switch ($module.Status) {
            'Fully Functional' { 'Green' }
            'Installed' { 'Yellow' }
            'Partially Functional' { 'Yellow' }
            'Import Failed' { 'Red' }
            'Not Installed' { 'Red' }
            'Error' { 'Red' }
            default { 'Gray' }
        }
        
        Write-Host "  $($module.ModuleName) - $($module.Status)" -ForegroundColor $statusColor
        if ($module.Installed) {
            Write-Host "    Version: $($module.Version)" -ForegroundColor Gray
            Write-Host "    Updated: $($module.LastUpdated)" -ForegroundColor Gray
        }
    }
}

# Test connectivity if requested
if ($TestConnectivity) {
    Write-Host "`n===========================================" -ForegroundColor Cyan
    Write-Host "CONNECTIVITY TESTING" -ForegroundColor Cyan
    Write-Host "===========================================" -ForegroundColor Cyan
    Write-Host "Note: This will attempt to connect to Microsoft 365 services" -ForegroundColor Yellow
    Write-Host "You may be prompted for authentication..." -ForegroundColor Yellow
    
    $connectivityResults = @()
    
    # Test Microsoft Graph
    if ($results | Where-Object { $_.ModuleName -eq 'Microsoft.Graph' -and $_.Status -eq 'Fully Functional' }) {
        Write-Host "`nTesting Microsoft Graph connectivity..." -ForegroundColor Cyan
        try {
            Connect-MgGraph -Scopes "User.Read" -NoWelcome -ErrorAction Stop
            $context = Get-MgContext
            Write-Host "✓ Microsoft Graph connected successfully" -ForegroundColor Green
            Write-Host "  Tenant: $($context.TenantId)" -ForegroundColor Gray
            Write-Host "  Account: $($context.Account)" -ForegroundColor Gray
            $connectivityResults += "Microsoft Graph: Connected"
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        }
        catch {
            Write-Host "✗ Microsoft Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
            $connectivityResults += "Microsoft Graph: Failed"
        }
    }
    
    # Add connectivity results to report
    foreach ($result in $results) {
        $result | Add-Member -MemberType NoteProperty -Name ConnectivityTest -Value ($connectivityResults -join '; ') -Force
    }
}

# Generate HTML report if requested
if ($GenerateReport) {
    Write-Host "`nGenerating HTML report..." -ForegroundColor Cyan
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Microsoft 365 PowerShell Modules Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2 { color: #0078d4; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #0078d4; color: white; }
        .status-good { background-color: #d4edda; }
        .status-warning { background-color: #fff3cd; }
        .status-error { background-color: #f8d7da; }
        .summary { background-color: #e3f2fd; padding: 15px; border-radius: 5px; margin: 20px 0; }
    </style>
</head>
<body>
    <h1>Microsoft 365 PowerShell Modules Report</h1>
    <div class="summary">
        <h2>Summary</h2>
        <p><strong>Generated:</strong> $(Get-Date)</p>
        <p><strong>Total Modules:</strong> $total</p>
        <p><strong>Installed:</strong> $installed</p>
        <p><strong>Fully Functional:</strong> $fullyFunctional</p>
    </div>
    
    <h2>Module Details</h2>
    <table>
        <tr>
            <th>Module Name</th>
            <th>Category</th>
            <th>Status</th>
            <th>Version</th>
            <th>Last Updated</th>
            <th>Commands Available</th>
            <th>Commands Missing</th>
        </tr>
"@

    foreach ($result in $results | Sort-Object Category, ModuleName) {
        $statusClass = switch ($result.Status) {
            'Fully Functional' { 'status-good' }
            'Installed' { 'status-warning' }
            'Partially Functional' { 'status-warning' }
            default { 'status-error' }
        }
        
        $commandsAvailable = ($result.CommandsAvailable -join ', ')
        $commandsMissing = ($result.CommandsMissing -join ', ')
        
        $html += @"
        <tr class="$statusClass">
            <td>$($result.ModuleName)</td>
            <td>$($result.Category)</td>
            <td>$($result.Status)</td>
            <td>$($result.Version)</td>
            <td>$($result.LastUpdated)</td>
            <td>$commandsAvailable</td>
            <td>$commandsMissing</td>
        </tr>
"@
    }
    
    $html += @"
    </table>
    
    <h2>Recommendations</h2>
    <ul>
"@

    # Add recommendations based on results
    $notInstalled = $results | Where-Object { -not $_.Installed }
    if ($notInstalled) {
        $html += "<li>Install missing modules: $($notInstalled.ModuleName -join ', ')</li>"
    }
    
    $importFailed = $results | Where-Object { $_.Status -eq 'Import Failed' }
    if ($importFailed) {
        $html += "<li>Troubleshoot import issues for: $($importFailed.ModuleName -join ', ')</li>"
    }
    
    $html += @"
        <li>Run regular updates with: Update-Module ModuleName</li>
        <li>Review the complete usage guide for detailed instructions</li>
    </ul>
</body>
</html>
"@

    try {
        $html | Out-File -FilePath $ReportPath -Encoding UTF8
        Write-Host "✓ HTML report generated: $ReportPath" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to generate report: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Recommendations
Write-Host "`n===========================================" -ForegroundColor Cyan
Write-Host "RECOMMENDATIONS" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan

$notInstalled = $results | Where-Object { -not $_.Installed }
if ($notInstalled.Count -gt 0) {
    Write-Host "Missing modules:" -ForegroundColor Yellow
    foreach ($module in $notInstalled) {
        Write-Host "  - $($module.ModuleName) ($($module.Description))" -ForegroundColor Red
    }
    Write-Host "Run: .\Install-Microsoft365Modules.ps1 -Force" -ForegroundColor Cyan
}

$importFailed = $results | Where-Object { $_.Status -eq 'Import Failed' }
if ($importFailed.Count -gt 0) {
    Write-Host "`nModules with import issues:" -ForegroundColor Yellow
    foreach ($module in $importFailed) {
        Write-Host "  - $($module.ModuleName)" -ForegroundColor Red
    }
    Write-Host "Try: Update-Module ModuleName -Force" -ForegroundColor Cyan
}

if ($installed -eq $total -and $fullyFunctional -eq $total) {
    Write-Host "`n🎉 All modules are installed and fully functional!" -ForegroundColor Green
}
else {
    Write-Host "`n⚠️ Some modules need attention. Review the details above." -ForegroundColor Yellow
}

Write-Host "`nVerification completed!" -ForegroundColor Cyan