# Microsoft 365 PowerShell Modules - Complete Installation and Usage Guide

## 📋 Overview

This comprehensive package provides enterprise-grade PowerShell module installation and management for Microsoft 365 services. It includes support for all major Microsoft 365 workloads with enhanced security, error handling, and logging capabilities.

## 🚀 Supported Services and Modules

### Core Microsoft 365 Services

| Service | Module | Description |
|---------|--------|-------------|
| **Microsoft Entra ID** | Microsoft.Graph | Unified API for identity and access management |
| **Exchange Online** | ExchangeOnlineManagement | Email administration and Defender for Office 365 |
| **SharePoint Online** | PnP.PowerShell<br>Microsoft.Online.SharePoint.PowerShell | Modern and legacy SharePoint management |
| **Microsoft Teams** | MicrosoftTeams | Teams administration and policy management |
| **Microsoft Defender for Office 365** | ExchangeOnlineManagement | Advanced threat protection and security |
| **Microsoft Purview Compliance** | Microsoft.Graph | Compliance and data governance (via Graph API) |

### Additional Platform Services

| Service | Module | Description |
|---------|--------|-------------|
| **Power Platform** | Microsoft.PowerApps.Administration.PowerShell<br>Microsoft.PowerApps.PowerShell | PowerApps and Power Automate management |
| **Microsoft Intune** | Microsoft.Graph.Intune | Mobile device and application management |
| **Dynamics 365** | Microsoft.Xrm.Data.PowerShell | Customer engagement platform |

### Legacy Support (Optional)

| Service | Module | Description |
|---------|--------|-------------|
| **Azure AD (Legacy)** | MSOnline<br>AzureAD | Legacy Azure Active Directory modules |

## 📁 Package Contents

```
c:\temp\
├── Install-Microsoft365Modules.ps1      # Main installation script (Enterprise)
├── Quick-Install-M365Modules.ps1        # Quick installation script
├── README-Microsoft365-PowerShell.md    # This guide
└── Microsoft365ModuleInstall_*.log      # Installation logs (generated)
```

## 🛠️ Installation Methods

### Method 1: Enterprise Installation (Recommended)

The main script provides comprehensive features for enterprise environments:

```powershell
# Run as Administrator for system-wide installation
.\Install-Microsoft365Modules.ps1

# Current user installation (no admin required)
.\Install-Microsoft365Modules.ps1 -Scope CurrentUser

# Force reinstall all modules
.\Install-Microsoft365Modules.ps1 -Force

# Include legacy Azure AD modules
.\Install-Microsoft365Modules.ps1 -IncludeLegacy

# Custom log location
.\Install-Microsoft365Modules.ps1 -LogPath "C:\Logs\M365Install.log"
```

**Enterprise Features:**
- ✅ Comprehensive prerequisite checking
- ✅ Detailed logging with categorization
- ✅ Retry logic with exponential backoff
- ✅ Module verification and testing
- ✅ Security validation (TLS 1.2, signatures)
- ✅ Progress tracking and timing
- ✅ Comprehensive error handling

### Method 2: Quick Installation

For immediate deployment and testing:

```powershell
.\Quick-Install-M365Modules.ps1
```

**Quick Features:**
- ✅ Streamlined installation process
- ✅ Progress tracking
- ✅ Basic verification
- ✅ User-friendly output

### Method 3: Manual Installation

Install specific modules individually:

```powershell
# Core modules
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser
Install-Module MicrosoftTeams -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser

# SharePoint Official Module
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser

# Power Platform
Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
Install-Module Microsoft.PowerApps.PowerShell -Scope CurrentUser

# Intune
Install-Module Microsoft.Graph.Intune -Scope CurrentUser

# Legacy (if needed)
Install-Module MSOnline -Scope CurrentUser
Install-Module AzureAD -Scope CurrentUser
```

## 🔐 Prerequisites

### System Requirements
- **PowerShell**: 5.1 or later (PowerShell 7+ recommended)
- **Operating System**: Windows 10/11, Windows Server 2016+
- **.NET Framework**: 4.7.2 or later
- **Execution Policy**: RemoteSigned or Unrestricted
- **Internet Connection**: Required for module downloads

### Permissions
- **AllUsers Scope**: Administrator privileges required
- **CurrentUser Scope**: Standard user permissions sufficient

### Security Requirements
- TLS 1.2 enabled (automatically configured)
- PSGallery repository access
- Valid Microsoft 365 tenant credentials

## 📚 Connection Examples and Usage

### Microsoft Graph (Entra ID & Microsoft 365)

```powershell
# Interactive authentication with specific scopes
Connect-MgGraph -Scopes @(
    "User.Read.All",
    "Group.Read.All", 
    "Directory.Read.All",
    "Application.Read.All"
)

# Verify connection
Get-MgContext

# Common operations
Get-MgUser -Top 10 | Select-Object DisplayName, UserPrincipalName, Id
Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -Top 10
Get-MgApplication -Top 10

# Compliance and security
Connect-MgGraph -Scopes "CompliancePolicy.Read.All", "SecurityEvents.Read.All"
Get-MgComplianceEdiscoveryCase
Get-MgSecurityAlert -Top 10
```

### SharePoint Online Management

#### Option 1: PnP PowerShell (Recommended)
```powershell
# Connect to SharePoint admin center
Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive

# Connect to specific site
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite" -Interactive

# Common operations
Get-PnPTenantSite
Get-PnPSite
Get-PnPList
Get-PnPListItem -List "Documents"

# Advanced operations
Set-PnPTenantSite -Identity "https://yourtenant.sharepoint.com/sites/site1" -SharingCapability ExternalUserSharingOnly
New-PnPList -Title "Custom List" -Template GenericList
```

#### Option 2: Official SharePoint Module
```powershell
# Connect to SharePoint admin center
Connect-SPOService -Url "https://yourtenant-admin.sharepoint.com"

# Common operations
Get-SPOSite
Get-SPOTenant
Set-SPOSite -Identity "https://yourtenant.sharepoint.com/sites/site1" -SharingCapability ExternalUserSharingOnly
```

### Microsoft Teams

```powershell
# Connect to Teams
Connect-MicrosoftTeams

# Verify connection
Get-CsTenant

# Common operations
Get-Team
Get-TeamChannel -GroupId (Get-Team -DisplayName "Sales Team").GroupId
Get-CsTeamsClientConfiguration
Get-CsTeamsMessagingPolicy

# Policy management
New-CsTeamsMessagingPolicy -Identity "RestrictedMessaging" -AllowUrlPreviews $false
Grant-CsTeamsMessagingPolicy -Identity "user@domain.com" -PolicyName "RestrictedMessaging"
```

### Exchange Online & Defender for Office 365

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Verify connection
Get-OrganizationConfig

# Exchange management
Get-Mailbox -ResultSize 10
Get-DistributionGroup
Get-TransportRule
Get-RetentionPolicy

# Defender for Office 365
Get-SafeLinksPolicy
Get-SafeAttachmentPolicy
Get-AntiPhishPolicy
Get-ATPPolicyForO365

# Compliance
Get-ComplianceSearchAction
Get-RetentionCompliancePolicy
Start-ComplianceSearch -Name "Legal Hold Search" -ContentMatchQuery "Subject:Contract"
```

### Power Platform Administration

```powershell
# Connect to Power Platform
Add-PowerAppsAccount

# Environment management
Get-PowerAppEnvironment
New-PowerAppEnvironment -DisplayName "Development" -LocationName "unitedstates"

# App management
Get-PowerApp
Get-PowerAppConnection
Get-PowerAutomate
```

### Microsoft Intune

```powershell
# Connect to Intune
Connect-MSGraph

# Device management
Get-IntuneManagedDevice
Get-IntuneApplication
Get-IntuneDeviceCompliancePolicy

# Application management
Get-IntuneApplicationAssignment
Get-IntuneMobileAppCategory
```

## 🔒 Security and Permissions Guide

### Microsoft Graph Scopes Reference

#### Identity and Access Management
```powershell
$IdentityScopes = @(
    "User.Read.All",           # Read all users
    "User.ReadWrite.All",      # Manage users
    "Group.Read.All",          # Read all groups
    "Group.ReadWrite.All",     # Manage groups
    "Directory.Read.All",      # Read directory
    "Directory.ReadWrite.All", # Manage directory
    "Application.Read.All",    # Read applications
    "Application.ReadWrite.All" # Manage applications
)
```

#### Security and Compliance
```powershell
$SecurityScopes = @(
    "SecurityEvents.Read.All",      # Read security events
    "ThreatIndicators.ReadWrite.OwnedBy", # Manage threat indicators
    "CompliancePolicy.Read.All",    # Read compliance policies
    "eDiscovery.Read.All",         # Read eDiscovery cases
    "InformationProtectionPolicy.Read" # Read IP policies
)
```

#### Teams and Communication
```powershell
$TeamsScopes = @(
    "Team.ReadBasic.All",      # Read basic team info
    "Team.Read.All",           # Read all team data
    "Channel.ReadBasic.All",   # Read basic channel info
    "Chat.Read.All"            # Read chat messages
)
```

### Role-Based Access Requirements

| Service | Required Roles |
|---------|----------------|
| **Entra ID** | Global Administrator, User Administrator, Groups Administrator |
| **Exchange Online** | Exchange Administrator, Global Administrator |
| **Defender for Office 365** | Security Administrator, Exchange Administrator |
| **SharePoint Online** | SharePoint Administrator, Global Administrator |
| **Microsoft Teams** | Teams Administrator, Global Administrator |
| **Power Platform** | Power Platform Administrator, Global Administrator |
| **Intune** | Intune Administrator, Global Administrator |
| **Compliance** | Compliance Administrator, Security Administrator |

### Best Practices for Secure Authentication

1. **Use Principle of Least Privilege**
   ```powershell
   # Request only necessary scopes
   Connect-MgGraph -Scopes "User.Read.All" # Instead of Directory.ReadWrite.All
   ```

2. **Use Interactive Authentication for User Sessions**
   ```powershell
   # Preferred for interactive sessions
   Connect-MgGraph -Scopes "User.Read.All" # Uses device code flow
   Connect-ExchangeOnline # Uses modern auth
   ```

3. **Use Service Principals for Automation**
   ```powershell
   # For automated scripts (Azure Automation, etc.)
   $credential = Get-Credential
   Connect-MgGraph -ClientSecretCredential $credential -TenantId "your-tenant-id"
   ```

4. **Always Disconnect Sessions**
   ```powershell
   # End of session cleanup
   Disconnect-MgGraph
   Disconnect-ExchangeOnline -Confirm:$false
   Disconnect-MicrosoftTeams
   ```

## 🚨 Troubleshooting Guide

### Common Installation Issues

#### 1. Execution Policy Error
```powershell
# Problem: Cannot load because running of scripts is disabled
# Solution:
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### 2. Module Not Found After Installation
```powershell
# Problem: Module installed but not available
# Solution:
Import-Module ModuleName -Force
Get-Module -ListAvailable ModuleName
```

#### 3. Version Conflicts
```powershell
# Problem: Multiple versions installed
# Solution:
Uninstall-Module ModuleName -AllVersions
Install-Module ModuleName -Force
```

#### 4. Authentication Failures
```powershell
# Problem: Cannot authenticate to services
# Solutions:

# Clear cached credentials
Disconnect-MgGraph
Disconnect-ExchangeOnline
[Microsoft.Graph.PowerShell.Authentication.GraphSession]::Reset()

# Check tenant and permissions
Get-MgContext
Get-MgUser -Top 1 # Test permissions
```

#### 5. PowerShellGet Issues
```powershell
# Problem: PowerShellGet module outdated
# Solution:
Install-Module PowerShellGet -Force -AllowClobber
Update-Module PowerShellGet
```

### Performance Optimization

#### 1. Module Loading Optimization
```powershell
# Load only specific module components
Import-Module Microsoft.Graph.Users -Force
Import-Module Microsoft.Graph.Groups -Force
# Instead of importing the entire Microsoft.Graph module
```

#### 2. Connection Reuse
```powershell
# Check existing connections before connecting
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "User.Read.All"
}
```

#### 3. Batch Operations
```powershell
# Process multiple items efficiently
$users = Get-MgUser -All
$users | ForEach-Object -Parallel {
    # Process each user in parallel
} -ThrottleLimit 10
```

### Diagnostic Commands

```powershell
# Check module versions
Get-Module -ListAvailable | Where-Object {
    $_.Name -match "Microsoft.Graph|PnP.PowerShell|MicrosoftTeams|ExchangeOnlineManagement"
} | Sort-Object Name, Version

# Check active connections
Get-MgContext
Get-ConnectionInformation
Get-CsTenant -ErrorAction SilentlyContinue

# Test module functionality
Test-ModuleManifest (Get-Module Microsoft.Graph -ListAvailable)[0].Path
Get-Command -Module Microsoft.Graph -CommandType Cmdlet | Measure-Object

# Check PowerShell environment
$PSVersionTable
Get-ExecutionPolicy -List
Get-PSRepository
```

## 📊 Module Management and Maintenance

### Regular Updates

```powershell
# Update all Microsoft 365 modules
$M365Modules = @(
    'Microsoft.Graph',
    'PnP.PowerShell', 
    'MicrosoftTeams',
    'ExchangeOnlineManagement',
    'Microsoft.Online.SharePoint.PowerShell',
    'Microsoft.PowerApps.Administration.PowerShell'
)

foreach ($Module in $M365Modules) {
    try {
        Write-Host "Updating $Module..." -ForegroundColor Cyan
        Update-Module $Module -Force
        Write-Host "✓ $Module updated successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to update $Module`: $($_.Exception.Message)" -ForegroundColor Red
    }
}
```

### Health Monitoring

```powershell
# Create module health check script
function Test-M365ModuleHealth {
    $modules = @{
        'Microsoft.Graph' = 'Connect-MgGraph'
        'PnP.PowerShell' = 'Connect-PnPOnline'
        'MicrosoftTeams' = 'Connect-MicrosoftTeams'
        'ExchangeOnlineManagement' = 'Connect-ExchangeOnline'
    }
    
    foreach ($module in $modules.Keys) {
        $command = $modules[$module]
        try {
            Import-Module $module -Force -ErrorAction Stop
            if (Get-Command $command -ErrorAction SilentlyContinue) {
                Write-Host "✓ $module - Healthy" -ForegroundColor Green
            } else {
                Write-Host "⚠ $module - Command missing: $command" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "✗ $module - Import failed" -ForegroundColor Red
        }
    }
}

# Run health check
Test-M365ModuleHealth
```

## 📋 Sample Automation Scripts

### Daily Admin Tasks

```powershell
# Daily Microsoft 365 administration script
param(
    [string]$TenantDomain = "yourtenant.onmicrosoft.com",
    [string]$LogPath = "c:\temp\DailyTasks_$(Get-Date -Format 'yyyyMMdd').log"
)

# Connect to services
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All"
Connect-ExchangeOnline

# Generate reports
$users = Get-MgUser -All | Where-Object { $_.AccountEnabled -eq $true }
$groups = Get-MgGroup -All
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Create summary report
$report = @{
    Date = Get-Date
    TotalUsers = $users.Count
    TotalGroups = $groups.Count
    TotalMailboxes = $mailboxes.Count
    NewUsersToday = ($users | Where-Object { $_.CreatedDateTime -gt (Get-Date).AddDays(-1) }).Count
}

$report | Export-Csv -Path $LogPath -NoTypeInformation
Write-Host "Daily report saved to: $LogPath" -ForegroundColor Green

# Disconnect
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false
```

### Bulk User Management

```powershell
# Bulk user operations example
function New-BulkUsers {
    param(
        [Parameter(Mandatory)]
        [string]$CsvPath,
        [string]$DefaultPassword = "TempPass123!"
    )
    
    $users = Import-Csv $CsvPath
    
    foreach ($user in $users) {
        try {
            $passwordProfile = @{
                Password = $DefaultPassword
                ForceChangePasswordNextSignIn = $true
            }
            
            $newUser = New-MgUser -DisplayName $user.DisplayName `
                                  -UserPrincipalName $user.UserPrincipalName `
                                  -MailNickname $user.MailNickname `
                                  -PasswordProfile $passwordProfile `
                                  -AccountEnabled:$true
            
            Write-Host "✓ Created user: $($user.DisplayName)" -ForegroundColor Green
        }
        catch {
            Write-Host "✗ Failed to create $($user.DisplayName): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
```

## 📞 Support and Resources

### Official Documentation Links

- [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/)
- [SharePoint PnP PowerShell](https://pnp.github.io/powershell/)
- [Teams PowerShell Overview](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-overview)
- [Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
- [Microsoft Purview PowerShell](https://learn.microsoft.com/en-us/microsoft-365/compliance/)
- [Power Platform PowerShell](https://learn.microsoft.com/en-us/power-platform/admin/powershell-getting-started)

### Community Resources

- [Microsoft 365 Community](https://aka.ms/m365pnp)
- [PowerShell Gallery](https://www.powershellgallery.com/)
- [Microsoft Tech Community](https://techcommunity.microsoft.com/)

### Getting Help

```powershell
# Get help for specific commands
Get-Help Connect-MgGraph -Full
Get-Help Connect-ExchangeOnline -Examples
Get-Help Connect-PnPOnline -Online

# List available commands in a module
Get-Command -Module Microsoft.Graph.Users
Get-Command -Module ExchangeOnlineManagement
```

---

## 📝 Version History

- **v2.0** (September 2025)
  - Added comprehensive module support
  - Enhanced error handling and logging
  - Added Power Platform and Intune modules
  - Improved security and authentication guidance
  - Added troubleshooting and maintenance sections

- **v1.0** (Initial Release)
  - Basic module installation
  - Core Microsoft 365 services support

---

**Created by:** Microsoft 365 Administration Team  
**Last Updated:** September 18, 2025  
**PowerShell Version:** 5.1+ (PowerShell 7+ recommended)