# Microsoft 365 Modules Verification and Testing Script
# This script verifies installation and tests basic functionality of all Microsoft 365 PowerShell modules

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$TestConnectivity,
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateReport,
    
    [Parameter(Mandatory = $false)]
    [string]$ReportPath = "c:\temp\M365ModulesReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    
    [Parameter(Mandatory = $false)]
    [switch]$OpenReport,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacy,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeBeta,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeAzure,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludePartner,
    
    [Parameter(Mandatory = $false)]
    [switch]$DetailedOutput
)

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "Microsoft 365 Modules Verification & Testing Script" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "Version: 2.0 | Updated: September 2025" -ForegroundColor Gray

# Define comprehensive modules to verify (matching Install-Microsoft365Modules.ps1)
$ModulesToVerify = @(
    @{
        Name = 'Microsoft.Graph'
        TestCommands = @('Connect-MgGraph', 'Get-MgContext', 'Get-MgUser', 'Get-MgGroup')
        Category = 'Core'
        Description = 'Microsoft Graph PowerShell SDK - Unified API for Microsoft 365, Entra ID, and Azure services'
        MinVersion = '2.0.0'
        Priority = 1
    },
    @{
        Name = 'PnP.PowerShell'
        TestCommands = @('Connect-PnPOnline', 'Get-PnPContext', 'Get-PnPSite', 'Get-PnPList')
        Category = 'SharePoint'
        Description = 'SharePoint PnP PowerShell - Modern SharePoint Online management and automation'
        MinVersion = '2.0.0'
        Priority = 2
    },
    @{
        Name = 'Microsoft.Online.SharePoint.PowerShell'
        TestCommands = @('Connect-SPOService', 'Get-SPOSite', 'Get-SPOTenant')
        Category = 'SharePoint'
        Description = 'SharePoint Online Management Shell - Official SharePoint Online administration'
        MinVersion = '16.0.0'
        Priority = 3
    },
    @{
        Name = 'MicrosoftTeams'
        TestCommands = @('Connect-MicrosoftTeams', 'Get-Team', 'Get-CsTeamsClientConfiguration')
        Category = 'Teams'
        Description = 'Microsoft Teams PowerShell Module - Teams administration and policy management'
        MinVersion = '5.0.0'
        Priority = 4
    },
    @{
        Name = 'ExchangeOnlineManagement'
        TestCommands = @('Connect-ExchangeOnline', 'Get-Mailbox', 'Get-OrganizationConfig', 'Get-SafeLinksPolicy')
        Category = 'Exchange'
        Description = 'Exchange Online PowerShell - Exchange, Defender for Office 365, and compliance management'
        MinVersion = '3.0.0'
        Priority = 5
    },
    @{
        Name = 'Microsoft.PowerApps.Administration.PowerShell'
        TestCommands = @('Add-PowerAppsAccount', 'Get-PowerAppEnvironment', 'Get-PowerApp')
        Category = 'PowerPlatform'
        Description = 'PowerApps Administration - Power Platform administration and governance'
        MinVersion = '2.0.0'
        Priority = 6
    },
    @{
        Name = 'Microsoft.PowerApps.PowerShell'
        TestCommands = @('Get-PowerApp', 'Get-PowerAppEnvironment')
        Category = 'PowerPlatform'
        Description = 'PowerApps PowerShell - Power Platform development and management'
        MinVersion = '1.0.0'
        Priority = 7
    },
    @{
        Name = 'MSOnline'
        TestCommands = @('Connect-MsolService', 'Get-MsolUser', 'Get-MsolCompanyInformation')
        Category = 'Legacy'
        Description = 'Azure Active Directory (Legacy) - Legacy Azure AD management (use with Microsoft.Graph)'
        MinVersion = '1.1.0'
        Priority = 8
        InstallCondition = { $IncludeLegacy }
    },
    @{
        Name = 'AzureAD'
        TestCommands = @('Connect-AzureAD', 'Get-AzureADUser', 'Get-AzureADTenantDetail')
        Category = 'Legacy'
        Description = 'Azure Active Directory V2 (Legacy) - Legacy Azure AD management (use with Microsoft.Graph)'
        MinVersion = '2.0.0'
        Priority = 9
        InstallCondition = { $IncludeLegacy }
    },
    @{
        Name = 'Microsoft.Graph.Intune'
        TestCommands = @('Connect-MSGraph', 'Get-IntuneManagedDevice', 'Get-IntuneApplication')
        Category = 'Intune'
        Description = 'Microsoft Intune PowerShell - Device management and mobile application management'
        MinVersion = '6.0.0'
        Priority = 10
    },
    @{
        Name = 'Microsoft.Xrm.Data.PowerShell'
        TestCommands = @('Get-CrmConnection', 'Get-CrmRecords')
        Category = 'Dynamics'
        Description = 'Dynamics 365 PowerShell - Customer engagement and operations management'
        MinVersion = '2.8.0'
        Priority = 11
    },
    @{
        Name = 'Microsoft.PowerApps.Checker.PowerShell'
        TestCommands = @('Get-PowerAppsCheckerRulesets', 'Invoke-PowerAppsChecker')
        Category = 'PowerPlatform'
        Description = 'Power Platform Solution Checker - Code analysis and best practices validation'
        MinVersion = '1.0.0'
        Priority = 12
    },
    @{
        Name = 'Microsoft.WinGet.Client'
        TestCommands = @('Find-WinGetPackage', 'Install-WinGetPackage', 'Get-WinGetPackage')
        Category = 'WindowsManagement'
        Description = 'Windows Package Manager Client - Software deployment and management'
        MinVersion = '1.0.0'
        Priority = 13
    },
    @{
        Name = 'Microsoft.Graph.Authentication'
        TestCommands = @('Connect-MgGraph', 'Get-MgContext', 'Disconnect-MgGraph')
        Category = 'Core'
        Description = 'Microsoft Graph Authentication - Enhanced authentication capabilities'
        MinVersion = '1.0.0'
        Priority = 14
    },
    @{
        Name = 'Microsoft.Graph.Beta'
        TestCommands = @('Connect-MgGraph', 'Get-MgBetaUser', 'Get-MgBetaGroup')
        Category = 'Core'
        Description = 'Microsoft Graph Beta - Preview APIs for latest Microsoft 365 features'
        MinVersion = '1.0.0'
        Priority = 15
        InstallCondition = { $IncludeBeta }
    },
    @{
        Name = 'Microsoft.Graph.DeviceManagement'
        TestCommands = @('Get-MgDeviceManagementManagedDevice', 'Get-MgDeviceManagementDeviceConfiguration')
        Category = 'Intune'
        Description = 'Microsoft Graph Device Management - Extended Intune and device management capabilities'
        MinVersion = '1.0.0'
        Priority = 16
    },
    @{
        Name = 'Microsoft.Graph.Identity.Governance'
        TestCommands = @('Get-MgIdentityGovernanceAccessReview', 'Get-MgIdentityGovernanceEntitlementManagement')
        Category = 'Governance'
        Description = 'Microsoft Graph Identity Governance - Access reviews, entitlement management, PIM'
        MinVersion = '1.0.0'
        Priority = 17
    },
    @{
        Name = 'Microsoft.Graph.Security'
        TestCommands = @('Get-MgSecurityIncident', 'Get-MgSecurityAlert')
        Category = 'Security'
        Description = 'Microsoft Graph Security - Security incidents, alerts, and threat protection'
        MinVersion = '1.0.0'
        Priority = 18
    },
    @{
        Name = 'Microsoft.Graph.Compliance'
        TestCommands = @('Get-MgComplianceEdiscoveryCase', 'Get-MgSecurityRetentionPolicy')
        Category = 'Compliance'
        Description = 'Microsoft Graph Compliance - Data governance, retention, and compliance policies'
        MinVersion = '1.0.0'
        Priority = 19
    },
    @{
        Name = 'Microsoft.Graph.Reports'
        TestCommands = @('Get-MgReportEmailActivityUserDetail', 'Get-MgReportOffice365ActiveUserDetail')
        Category = 'Reporting'
        Description = 'Microsoft Graph Reports - Usage analytics and reporting for Microsoft 365 services'
        MinVersion = '1.0.0'
        Priority = 20
    },
    @{
        Name = 'Microsoft.Graph.WindowsUpdates'
        TestCommands = @('Get-MgWindowsUpdatesDeployment', 'New-MgWindowsUpdatesDeployment')
        Category = 'WindowsManagement'
        Description = 'Microsoft Graph Windows Updates - Windows Update for Business deployment service'
        MinVersion = '1.0.0'
        Priority = 21
    },
    @{
        Name = 'Az.Accounts'
        TestCommands = @('Connect-AzAccount', 'Get-AzContext', 'Set-AzContext')
        Category = 'Azure'
        Description = 'Azure PowerShell Accounts - Authentication and context management for Azure'
        MinVersion = '2.0.0'
        Priority = 22
        InstallCondition = { $IncludeAzure }
    },
    @{
        Name = 'Az.Resources'
        TestCommands = @('Get-AzResourceGroup', 'Get-AzSubscription', 'New-AzResourceGroup')
        Category = 'Azure'
        Description = 'Azure PowerShell Resources - Resource group and subscription management'
        MinVersion = '6.0.0'
        Priority = 23
        InstallCondition = { $IncludeAzure }
    },
    @{
        Name = 'Microsoft.Graph.Calendar'
        TestCommands = @('Get-MgUserCalendar', 'Get-MgUserEvent', 'New-MgUserEvent')
        Category = 'Productivity'
        Description = 'Microsoft Graph Calendar - Calendar and scheduling management'
        MinVersion = '1.0.0'
        Priority = 24
    },
    @{
        Name = 'Microsoft.Graph.Files'
        TestCommands = @('Get-MgUserDrive', 'Get-MgDriveItem', 'Copy-MgDriveItem')
        Category = 'SharePoint'
        Description = 'Microsoft Graph Files - OneDrive and SharePoint file management'
        MinVersion = '1.0.0'
        Priority = 25
    },
    @{
        Name = 'Microsoft.Graph.Mail'
        TestCommands = @('Get-MgUserMessage', 'Send-MgUserMail', 'Get-MgUserMailFolder')
        Category = 'Exchange'
        Description = 'Microsoft Graph Mail - Email and message management'
        MinVersion = '1.0.0'
        Priority = 26
    },
    @{
        Name = 'Microsoft.Graph.People'
        TestCommands = @('Get-MgUserPerson', 'Get-MgUserPeople')
        Category = 'Social'
        Description = 'Microsoft Graph People - People and organizational relationships'
        MinVersion = '1.0.0'
        Priority = 27
    },
    @{
        Name = 'PartnerCenter'
        TestCommands = @('Connect-PartnerCenter', 'Get-PartnerCustomer', 'Get-PartnerCustomerSubscription')
        Category = 'Partner'
        Description = 'Partner Center PowerShell - CSP and partner management (for MSPs)'
        MinVersion = '3.0.0'
        Priority = 28
        InstallCondition = { $IncludePartner }
    }
)

# Enhanced logging function
function Write-VerificationLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Debug')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'Info' { 'White' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
        'Debug' { 'Gray' }
    }
    
    if ($DetailedOutput -or $Level -ne 'Debug') {
        Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
    }
}

# Filter modules based on conditions (same logic as installer)
$modulesToProcess = $ModulesToVerify | Where-Object {
    if ($_.ContainsKey('InstallCondition') -and $_.InstallCondition -is [scriptblock]) {
        return (& $_.InstallCondition)
    }
    return $true
} | Sort-Object Priority

$results = @()

Write-VerificationLog "Starting verification of $($modulesToProcess.Count) Microsoft 365 PowerShell modules..." -Level Info
if ($IncludeLegacy) { Write-VerificationLog "Including Legacy modules (MSOnline, AzureAD)" -Level Info }
if ($IncludeBeta) { Write-VerificationLog "Including Beta/Preview modules" -Level Info }
if ($IncludeAzure) { Write-VerificationLog "Including Azure PowerShell modules" -Level Info }
if ($IncludePartner) { Write-VerificationLog "Including Partner Center modules" -Level Info }

foreach ($moduleInfo in $modulesToProcess) {
    $moduleName = $moduleInfo.Name
    Write-VerificationLog "Checking $moduleName..." -Level Info
    
    $result = [PSCustomObject]@{
        ModuleName = $moduleName
        Category = $moduleInfo.Category
        Description = $moduleInfo.Description
        MinVersion = $moduleInfo.MinVersion
        Installed = $false
        Version = 'Not Installed'
        VersionStatus = 'Unknown'
        Location = 'N/A'
        CommandsAvailable = @()
        CommandsMissing = @()
        CommandsTotal = $moduleInfo.TestCommands.Count
        Status = 'Not Installed'
        ImportSuccess = $false
        LastUpdated = 'Unknown'
        Size = 'Unknown'
        Publisher = 'Unknown'
        InstallScope = 'Unknown'
        Priority = $moduleInfo.Priority
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
            $result.Publisher = $installedModule.Author
            
            # Check version against minimum requirement
            if ($moduleInfo.MinVersion) {
                try {
                    $currentVersion = [version]$installedModule.Version
                    $minimumVersion = [version]$moduleInfo.MinVersion
                    if ($currentVersion -ge $minimumVersion) {
                        $result.VersionStatus = 'Meets Requirements'
                    } else {
                        $result.VersionStatus = 'Below Minimum'
                        Write-VerificationLog "  ⚠️ Version $($result.Version) is below minimum required $($moduleInfo.MinVersion)" -Level Warning
                    }
                } catch {
                    $result.VersionStatus = 'Cannot Compare'
                }
            } else {
                $result.VersionStatus = 'No Requirement'
            }
            
            # Get installation scope
            try {
                $moduleScope = Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($moduleScope) {
                    $result.InstallScope = if ($moduleScope.InstalledLocation -like "*Program Files*") { "AllUsers" } else { "CurrentUser" }
                }
            } catch {
                $result.InstallScope = 'Unknown'
            }
            
            # Get module size
            try {
                $moduleSize = Get-ChildItem -Path $installedModule.ModuleBase -Recurse -File | 
                             Measure-Object -Property Length -Sum | 
                             Select-Object -ExpandProperty Sum
                $result.Size = "{0:N2} MB" -f ($moduleSize / 1MB)
            } catch {
                $result.Size = 'Unknown'
            }
            
            # Get last write time as update indicator
            try {
                $moduleManifest = Join-Path $installedModule.ModuleBase "$moduleName.psd1"
                if (Test-Path $moduleManifest) {
                    $lastWrite = (Get-Item $moduleManifest).LastWriteTime
                    $result.LastUpdated = $lastWrite.ToString('yyyy-MM-dd')
                    $daysSinceUpdate = (Get-Date - $lastWrite).Days
                    if ($daysSinceUpdate -gt 90) {
                        Write-VerificationLog "  ⚠️ Module is $daysSinceUpdate days old - consider updating" -Level Warning
                    }
                }
            }
            catch {
                $result.LastUpdated = 'Unknown'
            }
            
            Write-VerificationLog "  ✓ Installed: v$($result.Version) ($($result.VersionStatus))" -Level Success
            
            # Test module import
            try {
                Import-Module $moduleName -Force -ErrorAction Stop
                $result.ImportSuccess = $true
                Write-VerificationLog "  ✓ Import successful" -Level Success
                
                # Test commands
                foreach ($command in $moduleInfo.TestCommands) {
                    if (Get-Command $command -ErrorAction SilentlyContinue) {
                        $result.CommandsAvailable += $command
                        Write-VerificationLog "  ✓ Command available: $command" -Level Debug
                    }
                    else {
                        $result.CommandsMissing += $command
                        Write-VerificationLog "  ✗ Command missing: $command" -Level Warning
                    }
                }
                
                # Determine final status
                if ($result.CommandsMissing.Count -eq 0) {
                    $result.Status = 'Fully Functional'
                    Write-VerificationLog "  ✓ All $($result.CommandsTotal) commands available" -Level Success
                }
                elseif ($result.CommandsAvailable.Count -gt 0) {
                    $result.Status = 'Partially Functional'
                    Write-VerificationLog "  ⚠️ $($result.CommandsAvailable.Count)/$($result.CommandsTotal) commands available" -Level Warning
                }
                else {
                    $result.Status = 'Import Failed - No Commands'
                    Write-VerificationLog "  ✗ No expected commands found" -Level Error
                }
            }
            catch {
                $result.ImportSuccess = $false
                $result.Status = 'Import Failed'
                $result.CommandsMissing = $moduleInfo.TestCommands
                Write-VerificationLog "  ✗ Import failed: $($_.Exception.Message)" -Level Error
            }
        }
        else {
            Write-VerificationLog "  ✗ Not installed" -Level Warning
        }
    }
    catch {
        Write-VerificationLog "  ✗ Error checking module: $($_.Exception.Message)" -Level Error
        $result.Status = 'Error'
    }
    
    $results += $result
}

# Display summary
Write-Host "`n=====================================================" -ForegroundColor Cyan
Write-Host "VERIFICATION SUMMARY" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

$installed = ($results | Where-Object { $_.Installed }).Count
$fullyFunctional = ($results | Where-Object { $_.Status -eq 'Fully Functional' }).Count
$partiallyFunctional = ($results | Where-Object { $_.Status -eq 'Partially Functional' }).Count
$importFailed = ($results | Where-Object { $_.Status -like '*Import Failed*' }).Count
$notInstalled = ($results | Where-Object { -not $_.Installed }).Count
$belowMinimum = ($results | Where-Object { $_.VersionStatus -eq 'Below Minimum' }).Count
$total = $results.Count

Write-Host "Total modules checked: $total" -ForegroundColor White
Write-Host "✓ Installed modules: $installed" -ForegroundColor Green
Write-Host "✓ Fully functional: $fullyFunctional" -ForegroundColor Green
Write-Host "⚠️ Partially functional: $partiallyFunctional" -ForegroundColor Yellow
Write-Host "✗ Import failed: $importFailed" -ForegroundColor Red
Write-Host "✗ Not installed: $notInstalled" -ForegroundColor Red
if ($belowMinimum -gt 0) {
    Write-Host "⚠️ Below minimum version: $belowMinimum" -ForegroundColor Yellow
}

# Calculate overall health score
$healthScore = if ($total -gt 0) { 
    [math]::Round((($fullyFunctional * 100) / $total), 1) 
} else { 0 }

$healthColor = if ($healthScore -ge 90) { 'Green' } 
               elseif ($healthScore -ge 70) { 'Yellow' } 
               else { 'Red' }

Write-Host "`nOverall Health Score: $healthScore%" -ForegroundColor $healthColor

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
    
    # Add System.Web for HTML encoding
    Add-Type -AssemblyName System.Web
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Microsoft 365 PowerShell Modules Verification Report</title>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background-color: #f5f5f5; 
        }
        .container { 
            max-width: 1200px; 
            margin: 0 auto; 
            background-color: white; 
            border-radius: 8px; 
            box-shadow: 0 2px 10px rgba(0,0,0,0.1); 
            padding: 30px; 
        }
        .header { 
            text-align: center; 
            margin-bottom: 30px; 
            padding-bottom: 20px; 
            border-bottom: 2px solid #0078d4; 
        }
        .header h1 { 
            color: #0078d4; 
            margin: 0; 
            font-size: 2.2em; 
        }
        .header .subtitle { 
            color: #666; 
            margin: 10px 0 0 0; 
            font-size: 1.1em; 
        }
        .summary { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); 
            gap: 20px; 
            margin-bottom: 30px; 
        }
        .summary-card { 
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); 
            border-radius: 8px; 
            padding: 20px; 
            text-align: center; 
            border-left: 4px solid #0078d4; 
        }
        .summary-card h3 { 
            margin: 0 0 10px 0; 
            color: #333; 
            font-size: 2em; 
        }
        .summary-card p { 
            margin: 0; 
            color: #666; 
            font-weight: 500; 
        }
        .health-score { 
            text-align: center; 
            margin: 30px 0; 
            padding: 20px; 
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); 
            border-radius: 8px; 
            border: 2px solid #2196f3; 
        }
        .health-score h2 { 
            margin: 0; 
            font-size: 2.5em; 
            color: #1976d2; 
        }
        .table-container { 
            overflow-x: auto; 
            margin-top: 20px; 
        }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            background-color: white; 
            border-radius: 8px; 
            overflow: hidden; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.1); 
        }
        th { 
            background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%); 
            color: white; 
            padding: 15px 12px; 
            text-align: left; 
            font-weight: 600; 
            border-bottom: 2px solid #005a9e; 
        }
        td { 
            padding: 12px; 
            border-bottom: 1px solid #eee; 
            vertical-align: top; 
        }
        tr:nth-child(even) { 
            background-color: #f8f9fa; 
        }
        tr:hover { 
            background-color: #e3f2fd; 
            transition: background-color 0.2s; 
        }
        .status-fully-functional { 
            background-color: #d4edda; 
            color: #155724; 
            padding: 6px 12px; 
            border-radius: 20px; 
            font-weight: 500; 
            display: inline-block; 
            min-width: 80px; 
            text-align: center; 
        }
        .status-partially-functional { 
            background-color: #fff3cd; 
            color: #856404; 
            padding: 6px 12px; 
            border-radius: 20px; 
            font-weight: 500; 
            display: inline-block; 
            min-width: 80px; 
            text-align: center; 
        }
        .status-not-installed { 
            background-color: #f8d7da; 
            color: #721c24; 
            padding: 6px 12px; 
            border-radius: 20px; 
            font-weight: 500; 
            display: inline-block; 
            min-width: 80px; 
            text-align: center; 
        }
        .status-import-failed { 
            background-color: #f8d7da; 
            color: #721c24; 
            padding: 6px 12px; 
            border-radius: 20px; 
            font-weight: 500; 
            display: inline-block; 
            min-width: 80px; 
            text-align: center; 
        }
        .status-installed { 
            background-color: #d1ecf1; 
            color: #0c5460; 
            padding: 6px 12px; 
            border-radius: 20px; 
            font-weight: 500; 
            display: inline-block; 
            min-width: 80px; 
            text-align: center; 
        }
        .status-error { 
            background-color: #f8d7da; 
            color: #721c24; 
            padding: 6px 12px; 
            border-radius: 20px; 
            font-weight: 500; 
            display: inline-block; 
            min-width: 80px; 
            text-align: center; 
        }
        .status-unknown { 
            background-color: #e2e3e5; 
            color: #383d41; 
            padding: 6px 12px; 
            border-radius: 20px; 
            font-weight: 500; 
            display: inline-block; 
            min-width: 80px; 
            text-align: center; 
        }
        .version-meets-requirements { 
            background-color: #d4edda; 
            color: #155724; 
            padding: 4px 8px; 
            border-radius: 12px; 
            font-size: 0.9em; 
            font-weight: 500; 
        }
        .version-below-minimum { 
            background-color: #fff3cd; 
            color: #856404; 
            padding: 4px 8px; 
            border-radius: 12px; 
            font-size: 0.9em; 
            font-weight: 500; 
        }
        .version-no-requirement { 
            background-color: #e2e3e5; 
            color: #383d41; 
            padding: 4px 8px; 
            border-radius: 12px; 
            font-size: 0.9em; 
            font-weight: 500; 
        }
        .category-core { border-left: 4px solid #0078d4; }
        .category-sharepoint { border-left: 4px solid #00bcf2; }
        .category-teams { border-left: 4px solid #6264a7; }
        .category-exchange { border-left: 4px solid #0078d4; }
        .category-powerplatform { border-left: 4px solid #742774; }
        .category-intune { border-left: 4px solid #00bcf2; }
        .category-security { border-left: 4px solid #d83b01; }
        .category-compliance { border-left: 4px solid #107c10; }
        .category-governance { border-left: 4px solid #5c2d91; }
        .category-azure { border-left: 4px solid #0078d4; }
        .category-partner { border-left: 4px solid #ffb900; }
        .category-legacy { border-left: 4px solid #8a8886; }
        .category-dynamics { border-left: 4px solid #00bcf2; }
        .category-windowsmanagement { border-left: 4px solid #0078d4; }
        .category-reporting { border-left: 4px solid #107c10; }
        .category-productivity { border-left: 4px solid #5c2d91; }
        .category-social { border-left: 4px solid #ffb900; }
        .footer { 
            text-align: center; 
            margin-top: 30px; 
            padding-top: 20px; 
            border-top: 1px solid #ddd; 
            color: #666; 
            font-size: 0.9em; 
        }
        .recommendations { 
            background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%); 
            border: 1px solid #ffeb3b; 
            border-radius: 8px; 
            padding: 20px; 
            margin: 20px 0; 
        }
        .recommendations h3 { 
            color: #856404; 
            margin-top: 0; 
        }
        .recommendations ul { 
            margin: 10px 0; 
            padding-left: 20px; 
        }
        .recommendations li { 
            margin: 8px 0; 
            color: #856404; 
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Microsoft 365 PowerShell Modules</h1>
            <p class="subtitle">Verification Report - Generated on $(Get-Date -Format 'MMMM dd, yyyy at hh:mm tt')</p>
        </div>
        
        <div class="health-score">
            <h2>Overall Health Score: $healthScore%</h2>
        </div>
        
        <div class="summary">
            <div class="summary-card">
                <h3>$total</h3>
                <p>Total Modules</p>
            </div>
            <div class="summary-card">
                <h3>$installed</h3>
                <p>Installed</p>
            </div>
            <div class="summary-card">
                <h3>$fullyFunctional</h3>
                <p>Fully Functional</p>
            </div>
            <div class="summary-card">
                <h3>$partiallyFunctional</h3>
                <p>Partially Functional</p>
            </div>
            <div class="summary-card">
                <h3>$notInstalled</h3>
                <p>Not Installed</p>
            </div>
        </div>
"@

    # Add recommendations section if there are issues
    if ($notInstalled -gt 0 -or $importFailed -gt 0 -or $belowMinimum -gt 0) {
        $html += @"
        <div class="recommendations">
            <h3>🔧 Recommendations</h3>
            <ul>
"@
        if ($notInstalled -gt 0) {
            $html += "<li>Install missing modules using the Install-Microsoft365Modules.ps1 script</li>`n"
        }
        if ($importFailed -gt 0) {
            $html += "<li>Investigate import failures - check for conflicting modules or missing dependencies</li>`n"
        }
        if ($belowMinimum -gt 0) {
            $html += "<li>Update modules below minimum version requirements</li>`n"
        }
        $html += @"
                <li>Run 'Update-Module' to ensure all modules are up to date</li>
                <li>Consider using '-Force' parameter when importing modules to resolve conflicts</li>
            </ul>
        </div>
"@
    }

    $html += @"
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Module Name</th>
                        <th>Category</th>
                        <th>Status</th>
                        <th>Version</th>
                        <th>Version Status</th>
                        <th>Commands Available</th>
                        <th>Commands Missing</th>
                        <th>Last Updated</th>
                        <th>Size</th>
                    </tr>
                </thead>
                <tbody>
"@

    foreach ($result in $results | Sort-Object Category, ModuleName) {
        $statusClass = switch ($result.Status) {
            'Fully Functional' { 'status-fully-functional' }
            'Partially Functional' { 'status-partially-functional' }
            'Not Installed' { 'status-not-installed' }
            'Installed' { 'status-installed' }
            'Error' { 'status-error' }
            { $_ -like '*Import Failed*' } { 'status-import-failed' }
            default { 'status-unknown' }
        }
        
        $versionClass = switch ($result.VersionStatus) {
            'Meets Requirements' { 'version-meets-requirements' }
            'Below Minimum' { 'version-below-minimum' }
            default { 'version-no-requirement' }
        }
        
        $categoryClass = "category-$($result.Category.ToLower() -replace '[^a-z0-9]', '')"
        
        $commandsAvailable = if ($result.CommandsAvailable.Count -gt 0) { $result.CommandsAvailable.Count } else { "0" }
        $commandsMissing = if ($result.CommandsMissing.Count -gt 0) { $result.CommandsMissing.Count } else { "0" }
        
        # HTML encode function to prevent issues with special characters
        function Get-HtmlEncodedString($text) {
            if ([string]::IsNullOrEmpty($text)) { return "N/A" }
            return [System.Web.HttpUtility]::HtmlEncode($text.ToString())
        }
        
        $html += @"
                    <tr class="$categoryClass">
                        <td><strong>$(Get-HtmlEncodedString $result.ModuleName)</strong></td>
                        <td>$(Get-HtmlEncodedString $result.Category)</td>
                        <td><span class="$statusClass">$(Get-HtmlEncodedString $result.Status)</span></td>
                        <td>$(Get-HtmlEncodedString $result.Version)</td>
                        <td><span class="$versionClass">$(Get-HtmlEncodedString $result.VersionStatus)</span></td>
                        <td>$commandsAvailable / $($result.CommandsTotal)</td>
                        <td>$commandsMissing</td>
                        <td>$(Get-HtmlEncodedString $result.LastUpdated)</td>
                        <td>$(Get-HtmlEncodedString $result.Size)</td>
                    </tr>
"@
    }

    $html += @"
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p>Generated by Verify-M365Modules.ps1 | Microsoft 365 PowerShell Administration Tools</p>
            <p>For more information, visit <a href="https://docs.microsoft.com/en-us/powershell/module/" target="_blank">Microsoft PowerShell Documentation</a></p>
        </div>
    </div>
</body>
</html>
"@

    try {
        $html | Out-File -FilePath $ReportPath -Encoding UTF8 -Force
        Write-Host "✓ HTML report generated: $ReportPath" -ForegroundColor Green
        
        if ($OpenReport) {
            Write-Host "Opening report in default browser..." -ForegroundColor Cyan
            Start-Process $ReportPath
        }
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

$importFailed = $results | Where-Object { $_.Status -like '*Import Failed*' }
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

# Test connectivity if requested
if ($TestConnectivity) {
    Write-Host "`n===========================================" -ForegroundColor Cyan
    Write-Host "CONNECTIVITY TESTING" -ForegroundColor Cyan
    Write-Host "===========================================" -ForegroundColor Cyan
    
    $connectivityTests = @(
        @{
            Service = "Microsoft Graph"
            Module = "Microsoft.Graph.Authentication"
            TestCommand = "Get-MgContext"
            Description = "Microsoft Graph API connectivity"
        },
        @{
            Service = "Exchange Online"
            Module = "ExchangeOnlineManagement" 
            TestCommand = "Get-ConnectionInformation"
            Description = "Exchange Online PowerShell connectivity"
        },
        @{
            Service = "SharePoint Online"
            Module = "Microsoft.Online.SharePoint.PowerShell"
            TestCommand = "Get-SPOSite -Limit 1"
            Description = "SharePoint Online Management Shell connectivity"
        },
        @{
            Service = "Microsoft Teams"
            Module = "MicrosoftTeams"
            TestCommand = "Get-CsTeamsUpgradeStatus"
            Description = "Microsoft Teams PowerShell connectivity"
        },
        @{
            Service = "Azure Active Directory"
            Module = "AzureAD"
            TestCommand = "Get-AzureADCurrentSessionInfo"
            Description = "Azure AD PowerShell connectivity"
        },
        @{
            Service = "Microsoft Intune"
            Module = "Microsoft.Graph.Intune"
            TestCommand = "Get-IntuneManagedDevice -Top 1"
            Description = "Microsoft Intune Graph API connectivity"
        },
        @{
            Service = "Security & Compliance"
            Module = "ExchangeOnlineManagement"
            TestCommand = "Get-ComplianceSearch -ResultSize 1"
            Description = "Security & Compliance Center connectivity"
        },
        @{
            Service = "Power Platform"
            Module = "Microsoft.PowerApps.Administration.PowerShell"
            TestCommand = "Get-AdminPowerAppEnvironment -Limit 1"
            Description = "Power Platform administration connectivity"
        }
    )
    
    foreach ($test in $connectivityTests) {
        Write-Host "`nTesting $($test.Service)..." -ForegroundColor Yellow
        
        $moduleAvailable = Get-Module -ListAvailable -Name $test.Module -ErrorAction SilentlyContinue
        if (-not $moduleAvailable) {
            Write-Host "  ✗ Module $($test.Module) not installed" -ForegroundColor Red
            continue
        }
        
        try {
            $moduleImported = Get-Module -Name $test.Module -ErrorAction SilentlyContinue
            if (-not $moduleImported) {
                Import-Module $test.Module -Force -ErrorAction Stop
                Write-Host "  ✓ Module imported successfully" -ForegroundColor Green
            }
            
            $result = Invoke-Expression $test.TestCommand -ErrorAction Stop
            if ($result) {
                Write-Host "  ✓ Connectivity test passed" -ForegroundColor Green
                Write-Host "    $($test.Description)" -ForegroundColor Gray
            } else {
                Write-Host "  ⚠️ Connectivity test returned no data" -ForegroundColor Yellow
                Write-Host "    This may indicate authentication is required" -ForegroundColor Gray
            }
        }
        catch {
            Write-Host "  ✗ Connectivity test failed: $($_.Exception.Message)" -ForegroundColor Red
            if ($_.Exception.Message -like "*authentication*" -or $_.Exception.Message -like "*login*") {
                Write-Host "    Authentication may be required for this service" -ForegroundColor Gray
            }
        }
    }
    
    Write-Host "`nNote: Connectivity tests may require authentication." -ForegroundColor Cyan
    Write-Host "Use Connect-MgGraph, Connect-ExchangeOnline, Connect-SPOService, etc. to authenticate." -ForegroundColor Cyan
}

Write-Host "`nVerification completed!" -ForegroundColor Cyan