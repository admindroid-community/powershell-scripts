# Quick Microsoft 365 Modules Installation Script
# Simplified version for immediate deployment

Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "Microsoft 365 PowerShell Modules Installer" -ForegroundColor Cyan
Write-Host "Quick Installation Version" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan

# Enable TLS 1.2 for secure downloads
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Write-Host "✓ TLS 1.2 enabled for secure downloads" -ForegroundColor Green

# Install NuGet provider if needed
if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
    Write-Host "Installing NuGet provider..." -ForegroundColor Yellow
    Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -MinimumVersion 2.8.5.201
    Write-Host "✓ NuGet provider installed" -ForegroundColor Green
}

# Set PSGallery as trusted repository
if ((Get-PSRepository PSGallery).InstallationPolicy -ne 'Trusted') {
    Write-Host "Setting PSGallery as trusted repository..." -ForegroundColor Yellow
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Write-Host "✓ PSGallery configured as trusted" -ForegroundColor Green
}

# Define comprehensive modules list
$modules = @(
    @{ Name = 'Microsoft.Graph'; Description = 'Microsoft Graph SDK (Entra ID, Microsoft 365)' },
    @{ Name = 'PnP.PowerShell'; Description = 'SharePoint PnP PowerShell (Modern SharePoint)' },
    @{ Name = 'Microsoft.Online.SharePoint.PowerShell'; Description = 'SharePoint Online Management Shell (Official)' },
    @{ Name = 'MicrosoftTeams'; Description = 'Microsoft Teams PowerShell Module' },
    @{ Name = 'ExchangeOnlineManagement'; Description = 'Exchange Online & Defender for Office 365' },
    @{ Name = 'Microsoft.PowerApps.Administration.PowerShell'; Description = 'PowerApps Administration' },
    @{ Name = 'Microsoft.PowerApps.PowerShell'; Description = 'PowerApps PowerShell' },
    @{ Name = 'Microsoft.Graph.Intune'; Description = 'Microsoft Intune PowerShell' }
)

Write-Host "`nInstalling $($modules.Count) essential Microsoft 365 modules..." -ForegroundColor Cyan

$successCount = 0
$failedModules = @()

# Install each module with progress tracking
for ($i = 0; $i -lt $modules.Count; $i++) {
    $module = $modules[$i]
    $progress = [math]::Round((($i + 1) / $modules.Count) * 100)
    
    try {
        Write-Host "`n[$($i + 1)/$($modules.Count)] Installing $($module.Name)..." -ForegroundColor Cyan
        Write-Host "    $($module.Description)" -ForegroundColor Gray
        
        # Check if already installed
        $existing = Get-Module -ListAvailable -Name $module.Name -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "    Already installed: v$($existing[0].Version)" -ForegroundColor Yellow
        }
        
        # Install or update
        Install-Module -Name $module.Name -Scope CurrentUser -Force -AllowClobber -AcceptLicense -SkipPublisherCheck -ErrorAction Stop
        
        # Verify installation
        $installed = Get-Module -ListAvailable -Name $module.Name -ErrorAction SilentlyContinue
        if ($installed) {
            Write-Host "    ✓ Successfully installed: v$($installed[0].Version)" -ForegroundColor Green
            $successCount++
        }
        else {
            throw "Installation verification failed"
        }
    }
    catch {
        Write-Host "    ✗ Failed: $($_.Exception.Message)" -ForegroundColor Red
        $failedModules += $module.Name
    }
    
    # Progress indicator
    Write-Progress -Activity "Installing Microsoft 365 Modules" -Status "$progress% Complete" -PercentComplete $progress
}

# Clear progress bar
Write-Progress -Activity "Installing Microsoft 365 Modules" -Completed

# Installation summary
Write-Host "`n===========================================" -ForegroundColor Cyan
Write-Host "INSTALLATION SUMMARY" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "Successfully installed: $successCount of $($modules.Count) modules" -ForegroundColor Green

if ($failedModules.Count -gt 0) {
    Write-Host "Failed modules: $($failedModules -join ', ')" -ForegroundColor Red
    Write-Host "`nTo retry failed modules, run the main installation script:" -ForegroundColor Yellow
    Write-Host ".\Install-Microsoft365Modules.ps1 -Force" -ForegroundColor Yellow
}

# Quick verification
Write-Host "`nPerforming quick verification..." -ForegroundColor Cyan
$verificationCommands = @{
    'Microsoft.Graph' = 'Connect-MgGraph'
    'PnP.PowerShell' = 'Connect-PnPOnline'
    'Microsoft.Online.SharePoint.PowerShell' = 'Connect-SPOService'
    'MicrosoftTeams' = 'Connect-MicrosoftTeams'
    'ExchangeOnlineManagement' = 'Connect-ExchangeOnline'
    'Microsoft.PowerApps.Administration.PowerShell' = 'Add-PowerAppsAccount'
}

foreach ($moduleName in $verificationCommands.Keys) {
    $command = $verificationCommands[$moduleName]
    try {
        Import-Module $moduleName -Force -ErrorAction SilentlyContinue
        if (Get-Command $command -ErrorAction SilentlyContinue) {
            Write-Host "✓ $moduleName - Command '$command' available" -ForegroundColor Green
        }
        else {
            Write-Host "⚠ $moduleName - Command '$command' not found" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "⚠ $moduleName - Import failed" -ForegroundColor Yellow
    }
}

# Next steps
Write-Host "`n===========================================" -ForegroundColor Cyan
Write-Host "NEXT STEPS" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "1. Test connectivity to your Microsoft 365 tenant" -ForegroundColor White
Write-Host "2. Review the complete usage guide: .\README-Microsoft365-PowerShell.md" -ForegroundColor White
Write-Host "3. For detailed installation options, use: .\Install-Microsoft365Modules.ps1" -ForegroundColor White

Write-Host "`nQuick connection examples:" -ForegroundColor Yellow
Write-Host "Connect-MgGraph -Scopes 'User.Read.All'" -ForegroundColor Gray
Write-Host "Connect-ExchangeOnline" -ForegroundColor Gray
Write-Host "Connect-MicrosoftTeams" -ForegroundColor Gray

Write-Host "`nInstallation completed successfully! 🎉" -ForegroundColor Green