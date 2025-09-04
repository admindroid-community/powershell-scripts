<#
.SYNOPSIS
    Module Compatibility Fixer for M365 External User Audit
    
.DESCRIPTION
    This script fixes module compatibility issues by reinstalling/updating problematic modules
#>

[CmdletBinding()]
param()

Write-Host "🔧 M365 Module Compatibility Fixer" -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor Cyan

try {
    # Uninstall potentially conflicting versions
    Write-Host "`n1. Checking for conflicting module versions..." -ForegroundColor Yellow
    
    $modulesToFix = @(
        "Microsoft.Graph",
        "Microsoft.Graph.Authentication", 
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Groups"
    )
    
    foreach ($module in $modulesToFix) {
        $installedVersions = Get-Module -Name $module -ListAvailable
        if ($installedVersions.Count -gt 1) {
            Write-Host "⚠️  Multiple versions of $module found. Cleaning up..." -ForegroundColor Yellow
            
            # Keep only the latest version
            $latestVersion = $installedVersions | Sort-Object Version -Descending | Select-Object -First 1
            $oldVersions = $installedVersions | Where-Object { $_.Version -ne $latestVersion.Version }
            
            foreach ($oldVersion in $oldVersions) {
                try {
                    Write-Host "  Removing old version: $($oldVersion.Version)" -ForegroundColor Gray
                    Uninstall-Module -Name $module -RequiredVersion $oldVersion.Version -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host "  Could not remove $($oldVersion.Version): $($_.Exception.Message)" -ForegroundColor Gray
                }
            }
        }
    }
    
    # Reinstall Microsoft Graph with consistent versions
    Write-Host "`n2. Reinstalling Microsoft Graph modules..." -ForegroundColor Yellow
    
    try {
        # Uninstall all Graph modules first
        Write-Host "  Uninstalling existing Graph modules..." -ForegroundColor Gray
        $graphModules = Get-Module Microsoft.Graph* -ListAvailable
        foreach ($module in $graphModules) {
            try {
                Uninstall-Module -Name $module.Name -AllVersions -Force -ErrorAction SilentlyContinue
            }
            catch {
                # Ignore errors during uninstall
            }
        }
        
        # Install fresh versions
        Write-Host "  Installing fresh Microsoft Graph modules..." -ForegroundColor Gray
        Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
        
        Write-Host "✅ Microsoft Graph modules reinstalled successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Error reinstalling Graph modules: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Check PnP PowerShell
    Write-Host "`n3. Checking PnP PowerShell..." -ForegroundColor Yellow
    
    try {
        $pnpModule = Get-Module -Name "PnP.PowerShell" -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $pnpModule) {
            Write-Host "  Installing PnP PowerShell..." -ForegroundColor Gray
            Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
        } else {
            Write-Host "✅ PnP PowerShell is available: v$($pnpModule.Version)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "❌ Error with PnP PowerShell: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Check Exchange Online Management
    Write-Host "`n4. Checking Exchange Online Management..." -ForegroundColor Yellow
    
    try {
        $exoModule = Get-Module -Name "ExchangeOnlineManagement" -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $exoModule) {
            Write-Host "  Installing Exchange Online Management..." -ForegroundColor Gray
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
        } else {
            Write-Host "✅ Exchange Online Management is available: v$($exoModule.Version)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "❌ Error with Exchange Online Management: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`n✅ Module compatibility fix completed!" -ForegroundColor Green
    Write-Host "`n📝 Recommendations:" -ForegroundColor Cyan
    Write-Host "  1. Restart PowerShell session before running the audit" -ForegroundColor White
    Write-Host "  2. If issues persist, try running: Import-Module Microsoft.Graph -Force" -ForegroundColor White
    Write-Host "  3. Consider using PowerShell 7.x for better module compatibility" -ForegroundColor White
}
catch {
    Write-Host "`n❌ Module fix failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "You may need to manually reinstall the modules or run PowerShell as Administrator" -ForegroundColor Yellow
}
