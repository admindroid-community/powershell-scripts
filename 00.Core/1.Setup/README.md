# Microsoft 365 PowerShell Modules Installation Script

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

A comprehensive, enterprise-grade PowerShell script for installing and managing all Microsoft 365 administration modules. This script provides a one-stop solution for setting up a complete Microsoft 365 management environment with enhanced security, error handling, and modular installation options.

## 🚀 Features

- **Comprehensive Module Coverage**: Installs 28+ PowerShell modules covering all Microsoft 365 services
- **Enterprise-Grade Reliability**: Advanced error handling, retry logic, and comprehensive logging
- **Modular Installation**: Optional module categories to prevent bloat
- **Security First**: TLS 1.2 enforcement, trusted repositories, and secure authentication
- **Cross-Platform Ready**: Supports PowerShell 5.1+ and PowerShell 7+
- **Detailed Reporting**: Comprehensive installation summary and usage examples
- **Prerequisite Validation**: Automatic system requirements checking

## 📋 Supported Microsoft 365 Services

### Core Services
- **Microsoft Graph** - Unified API for Microsoft 365, Entra ID, and Azure services
- **Microsoft Entra ID (Azure AD)** - User, group, and directory management
- **Exchange Online** - Email, mailboxes, and Defender for Office 365
- **SharePoint Online** - Sites, lists, and content management (PnP and official modules)
- **Microsoft Teams** - Teams administration and policy management

### Extended Services
- **Microsoft Intune** - Device and application management
- **Power Platform** - PowerApps, Power Automate, and governance
- **Microsoft Purview** - Compliance, security, and data governance
- **Dynamics 365** - Customer engagement and operations management
- **Microsoft Graph Extended** - Security, Reports, Calendar, Files, Mail, People

### Optional Categories
- **Azure PowerShell** - Hybrid cloud resource management (`-IncludeAzure`)
- **Partner Center** - CSP and MSP customer management (`-IncludePartner`)
- **Beta/Preview** - Latest preview features (`-IncludeBeta`)
- **Legacy Modules** - MSOnline and AzureAD compatibility (`-IncludeLegacy`)

## 🛠️ Installation

### Prerequisites

- **Operating System**: Windows 10/11, Windows Server 2016+
- **PowerShell**: Version 5.1 or later (PowerShell 7+ recommended)
- **Execution Policy**: RemoteSigned or Unrestricted
- **Permissions**: Administrator rights (for AllUsers scope) or standard user (for CurrentUser scope)
- **.NET Framework**: 4.7.2 or later (recommended)
- **Internet Connection**: Required for module downloads from PowerShell Gallery

### Quick Start

1. **Download the script**:
   ```powershell
   # Clone the repository or download Install-Microsoft365Modules.ps1
   ```

2. **Run with default settings** (requires Administrator):
   ```powershell
   .\Install-Microsoft365Modules.ps1
   ```

3. **Run for current user only** (no Administrator required):
   ```powershell
   .\Install-Microsoft365Modules.ps1 -Scope CurrentUser
   ```

## 📖 Usage Examples

### Basic Installation
```powershell
# Install core Microsoft 365 modules (requires Administrator)
.\Install-Microsoft365Modules.ps1
```

### Current User Installation
```powershell
# Install for current user only (no Administrator required)
.\Install-Microsoft365Modules.ps1 -Scope CurrentUser
```

### Full Enterprise Installation
```powershell
# Install all modules including preview, Azure, and partner modules
.\Install-Microsoft365Modules.ps1 -IncludeBeta -IncludeAzure -IncludePartner -IncludeLegacy
```

### Force Reinstall
```powershell
# Force reinstall all modules (useful for updates)
.\Install-Microsoft365Modules.ps1 -Force
```

### Custom Log Location
```powershell
# Specify custom log file location
.\Install-Microsoft365Modules.ps1 -LogPath "D:\Logs\M365Install.log"
```

### CSP/MSP Environment
```powershell
# Installation optimized for CSP/MSP partners
.\Install-Microsoft365Modules.ps1 -IncludePartner -IncludeAzure -Scope CurrentUser
```

## 🔧 Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Force` | Switch | False | Forces reinstallation of modules even if they already exist |
| `-Scope` | String | 'AllUsers' | Installation scope: 'AllUsers' (requires admin) or 'CurrentUser' |
| `-LogPath` | String | Auto-generated | Path for log file (default: c:\temp\Microsoft365ModuleInstall_timestamp.log) |
| `-IncludeLegacy` | Switch | False | Include legacy Azure AD modules (MSOnline, AzureAD) |
| `-IncludeBeta` | Switch | False | Include Microsoft Graph Beta modules for preview features |
| `-IncludeAzure` | Switch | False | Include Azure PowerShell modules for hybrid management |
| `-IncludePartner` | Switch | False | Include Partner Center modules for CSP/MSP management |

## 📦 Installed Modules

<details>
<summary><strong>Core Modules (Always Installed)</strong></summary>

| Module | Description | Key Commands |
|--------|-------------|--------------|
| `Microsoft.Graph` | Unified API for Microsoft 365 and Azure services | `Connect-MgGraph`, `Get-MgUser`, `Get-MgGroup` |
| `Microsoft.Graph.Authentication` | Enhanced authentication capabilities | `Connect-MgGraph`, `Get-MgContext`, `Disconnect-MgGraph` |
| `PnP.PowerShell` | Modern SharePoint Online management | `Connect-PnPOnline`, `Get-PnPSite`, `Get-PnPList` |
| `Microsoft.Online.SharePoint.PowerShell` | Official SharePoint Online administration | `Connect-SPOService`, `Get-SPOSite`, `Get-SPOTenant` |
| `MicrosoftTeams` | Teams administration and policy management | `Connect-MicrosoftTeams`, `Get-Team`, `Get-CsTeamsClientConfiguration` |
| `ExchangeOnlineManagement` | Exchange and Defender for Office 365 | `Connect-ExchangeOnline`, `Get-Mailbox`, `Get-SafeLinksPolicy` |

</details>

<details>
<summary><strong>Extended Microsoft Graph Modules</strong></summary>

| Module | Description | Category |
|--------|-------------|----------|
| `Microsoft.Graph.DeviceManagement` | Extended Intune and device management | Intune |
| `Microsoft.Graph.Identity.Governance` | Access reviews, entitlement management, PIM | Governance |
| `Microsoft.Graph.Security` | Security incidents, alerts, threat protection | Security |
| `Microsoft.Graph.Compliance` | Data governance, retention policies | Compliance |
| `Microsoft.Graph.Reports` | Usage analytics and reporting | Reporting |
| `Microsoft.Graph.Calendar` | Calendar and scheduling management | Productivity |
| `Microsoft.Graph.Files` | OneDrive and SharePoint file management | SharePoint |
| `Microsoft.Graph.Mail` | Email and message management | Exchange |
| `Microsoft.Graph.People` | People and organizational relationships | Social |
| `Microsoft.Graph.WindowsUpdates` | Windows Update for Business | Windows Management |

</details>

<details>
<summary><strong>Optional Modules</strong></summary>

| Module | Parameter Required | Description |
|--------|--------------------|-------------|
| `Microsoft.Graph.Beta` | `-IncludeBeta` | Preview APIs for latest features |
| `Az.Accounts` | `-IncludeAzure` | Azure authentication and context |
| `Az.Resources` | `-IncludeAzure` | Azure resource management |
| `PartnerCenter` | `-IncludePartner` | CSP and partner management |
| `MSOnline` | `-IncludeLegacy` | Legacy Azure AD management |
| `AzureAD` | `-IncludeLegacy` | Legacy Azure AD V2 management |

</details>

## 🔑 Quick Start Commands

### Authentication Examples
```powershell
# Microsoft Graph (recommended for most scenarios)
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All"

# SharePoint Online (PnP - recommended)
Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive

# Exchange Online
Connect-ExchangeOnline

# Microsoft Teams
Connect-MicrosoftTeams

# Power Platform
Add-PowerAppsAccount
```

### Common Administration Tasks
```powershell
# Get all users
Get-MgUser -All

# Get all SharePoint sites
Get-PnPTenantSite

# Get all Teams
Get-Team

# Get all mailboxes
Get-Mailbox -ResultSize Unlimited

# Get security incidents
Connect-MgGraph -Scopes "SecurityIncident.Read.All"
Get-MgSecurityIncident

# Generate usage reports
Connect-MgGraph -Scopes "Reports.Read.All"
Get-MgReportOffice365ActiveUserDetail -Period D30
```

## 🔒 Security Best Practices

The script implements several security measures:

- **TLS 1.2 Enforcement**: Ensures secure downloads
- **Trusted Repository Validation**: Configures PSGallery as trusted
- **Module Signature Verification**: Validates publisher signatures
- **Least Privilege Access**: Recommends minimal required permissions
- **Secure Authentication**: Supports modern authentication methods

### Recommended Security Practices

1. **Use Interactive Authentication**: Prefer interactive login over stored credentials
2. **Implement Least Privilege**: Only request necessary scopes and permissions
3. **Regular Updates**: Keep modules updated with `Update-Module`
4. **Session Management**: Always disconnect sessions when finished
5. **Secure Credential Storage**: Never hardcode credentials in scripts

## 📊 Logging and Monitoring

The script provides comprehensive logging:

- **Detailed Progress Tracking**: Real-time installation progress
- **Error Reporting**: Comprehensive error details and retry information
- **Performance Metrics**: Installation duration and success rates
- **Security Events**: Authentication and permission validation logs
- **Module Verification**: Command availability and version validation

### Log File Example
```
[2025-09-23 14:30:15] [Info] [Main] MICROSOFT 365 POWERSHELL MODULES INSTALLATION SCRIPT v2.0
[2025-09-23 14:30:16] [Success] [Prerequisites] PowerShell version check passed: 7.4.5
[2025-09-23 14:30:17] [Success] [Security] TLS 1.2 security protocol confirmed
[2025-09-23 14:30:20] [Success] [Core] Microsoft.Graph v2.10.0 installed successfully
```

## 🔧 Troubleshooting

### Common Issues and Solutions

**Issue**: "Execution policy restricts script execution"
```powershell
# Solution: Update execution policy
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Issue**: "Administrator privileges required"
```powershell
# Solution: Use CurrentUser scope
.\Install-Microsoft365Modules.ps1 -Scope CurrentUser
```

**Issue**: "Module installation fails with network errors"
```powershell
# Solution: Check internet connection and retry
.\Install-Microsoft365Modules.ps1 -Force
```

**Issue**: "PowerShellGet is outdated"
```powershell
# Solution: Update PowerShellGet manually
Install-Module PowerShellGet -Force -AllowClobber
```

### Getting Help

1. **Check the log file** for detailed error information
2. **Run with verbose output**: Use PowerShell's `-Verbose` parameter
3. **Verify prerequisites**: Ensure all system requirements are met
4. **Test network connectivity**: Ensure access to PowerShell Gallery

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Test your changes thoroughly
4. Submit a pull request with detailed description

### Development Guidelines

- Follow PowerShell best practices
- Include comprehensive error handling
- Add appropriate logging statements
- Update documentation for new features
- Test on multiple PowerShell versions

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 📞 Support

For support and questions:

- **Issues**: Use GitHub Issues for bug reports and feature requests
- **Documentation**: Check the comprehensive inline help with `Get-Help .\Install-Microsoft365Modules.ps1 -Full`
- **Community**: Join the PowerShell community forums

## 🔄 Version History

### Version 2.0 (September 2025)
- Added 18 new Microsoft Graph modules
- Implemented modular installation with optional categories
- Enhanced security and error handling
- Added support for preview/beta modules
- Improved logging and reporting

### Version 1.0 (Initial Release)
- Core Microsoft 365 modules installation
- Basic error handling and logging
- Support for legacy modules

## 🎯 Roadmap

- [ ] Support for Government Cloud (GCC, GCC High, DoD)
- [ ] Automated module updates scheduling
- [ ] Configuration file support
- [ ] Integration with Azure DevOps pipelines
- [ ] Module dependency resolution
- [ ] Performance optimization for large environments

---

**Note**: This script is designed for administrative users and requires appropriate permissions. Always test in a non-production environment first.

For the most up-to-date information and latest features, please refer to the [official repository](https://github.com/your-username/admin-powershell-scripts).