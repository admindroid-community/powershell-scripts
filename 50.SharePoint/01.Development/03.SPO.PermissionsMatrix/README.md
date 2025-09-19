# SharePoint Online Permissions Matrix Audit

## Overview

The **SharePoint Online Permissions Matrix Audit** script provides comprehensive auditing of permissions across your entire SharePoint Online tenancy. This advanced PowerShell script generates detailed reports showing who has access to what resources, helping organizations maintain security compliance and governance.

**Author:** Alpesh Nakar  
**Company:** EDUC4TE  
**Website:** educ4te.com

## Key Features

### 🔍 **Comprehensive Audit Capabilities**
- **Site Collection Permissions**: Full audit of all site collection permissions
- **Subsite Permissions**: Optional deep-dive into subsite permissions
- **List/Library Permissions**: Detailed permissions for lists and document libraries
- **Sharing Links Audit**: Analysis of sharing links and external access
- **User & Group Permissions**: Complete breakdown of user and group access
- **Permission Inheritance**: Identifies inherited vs. direct permissions

### ⚡ **Performance & Throttling**
- **Configurable Throttling**: Process limited sites simultaneously (default: 2)
- **Batch Processing**: Efficient batch processing with automatic delays
- **Progress Tracking**: Real-time progress with percentage completion
- **Memory Efficient**: Optimized for large tenancies
- **API Rate Limiting**: Built-in protection against SharePoint throttling

### 🔐 **Authentication Methods**
- **Interactive Authentication**: Modern authentication with MFA support
- **Username/Password**: Traditional credential-based authentication
- **Certificate Authentication**: Secure app-only authentication for automation
- **Scheduler Friendly**: Support for automated execution

### 📊 **Advanced Reporting**
- **CSV Export**: Machine-readable format for analysis and compliance
- **HTML Report**: Professional formatted report with visual styling
- **Executive Summary**: High-level statistics and overview
- **Detailed Breakdowns**: Site-by-site permissions analysis
- **Error Reporting**: Comprehensive error tracking and reporting

### 🛠️ **Additional Features**
- **Site Filtering**: Target specific sites using keyword filters
- **Verbose Logging**: Detailed execution logging for troubleshooting
- **Module Validation**: Automatic validation of required PowerShell modules
- **Error Handling**: Robust error handling with detailed logging
- **Connection Management**: Automatic connection cleanup

## Prerequisites

### Required PowerShell Modules
- **PnP.PowerShell** (>= 1.12.0) - Primary SharePoint Online operations
- **Microsoft.Online.SharePoint.PowerShell** (>= 16.0.0) - SharePoint Online Management Shell
- **MSOnline** (>= 1.0.0) - Optional, for additional authentication scenarios

### Installation
```powershell
# Install required modules
Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -Scope CurrentUser

# Optional: Install MSOnline for additional authentication options
Install-Module -Name MSOnline -Force -Scope CurrentUser
```

### Module Import
The script automatically checks for and imports required modules. If modules are missing, you'll see clear installation instructions.

### Permissions Required
- **SharePoint Administrator** or **Global Administrator** role
- **Sites.Read.All** application permission (for app-only authentication)
- **User.Read.All** application permission (for user information)

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `UserName` | String | - | Username for authentication |
| `Password` | String | - | Password for authentication |
| `TenantName` | String | - | SharePoint tenant name (e.g., 'contoso') |
| `ClientId` | String | - | Azure AD App Registration Client ID |
| `CertificateThumbprint` | String | - | Certificate thumbprint for app-only auth |
| `ThrottleLimit` | Int | 2 | Number of sites to process simultaneously |
| `IncludeSubsites` | Switch | False | Include subsite permissions in audit |
| `IncludeListPermissions` | Switch | False | Include list/library permissions |
| `IncludeSharingLinks` | Switch | False | Include sharing links analysis |
| `GenerateHtmlReport` | Switch | False | Generate HTML formatted report |
| `OutputPath` | String | C:\temp | Output directory for reports |
| `SiteFilter` | Array | @() | Filter sites by keywords |
| `VerboseLogging` | Switch | False | Enable detailed logging |

## Usage Examples

### Basic Audit
```powershell
# Simple audit with interactive authentication
.\AuditSPO.PermissionsMatrix.ps1 -TenantName "contoso" -GenerateHtmlReport
```

### Comprehensive Audit
```powershell
# Full audit with all features enabled
.\AuditSPO.PermissionsMatrix.ps1 -TenantName "contoso" `
    -ThrottleLimit 3 `
    -IncludeSubsites `
    -IncludeListPermissions `
    -IncludeSharingLinks `
    -GenerateHtmlReport `
    -VerboseLogging `
    -OutputPath "C:\Reports"
```

### Credential-Based Authentication
```powershell
# Audit with username/password (suitable for scheduling)
.\AuditSPO.PermissionsMatrix.ps1 `
    -UserName "admin@contoso.onmicrosoft.com" `
    -Password "SecurePassword123" `
    -TenantName "contoso" `
    -GenerateHtmlReport `
    -OutputPath "C:\SharePointReports"
```

### Certificate-Based Authentication
```powershell
# Secure app-only authentication for production environments
.\AuditSPO.PermissionsMatrix.ps1 `
    -TenantName "contoso" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -CertificateThumbprint "1234567890ABCDEF1234567890ABCDEF12345678" `
    -GenerateHtmlReport `
    -ThrottleLimit 5
```

### Filtered Audit
```powershell
# Audit specific sites containing keywords
.\AuditSPO.PermissionsMatrix.ps1 -TenantName "contoso" `
    -SiteFilter @("finance", "hr", "legal") `
    -IncludeListPermissions `
    -GenerateHtmlReport
```

### Production Scheduled Execution
```powershell
# Enterprise-grade execution with all security features
.\AuditSPO.PermissionsMatrix.ps1 `
    -ClientId $env:SP_CLIENT_ID `
    -CertificateThumbprint $env:SP_CERT_THUMBPRINT `
    -TenantName $env:SP_TENANT_NAME `
    -ThrottleLimit 4 `
    -IncludeSubsites `
    -IncludeListPermissions `
    -GenerateHtmlReport `
    -OutputPath "\\server\reports\SharePoint" `
    -VerboseLogging
```

## Output Reports

### CSV Report
- **Filename**: `SPO_PermissionsMatrix_YYYYMMDD_HHMMSS.csv`
- **Format**: Machine-readable CSV for analysis
- **Content**: Detailed permissions matrix with all audit data
- **Use Case**: Data analysis, compliance reporting, PowerBI integration

### HTML Report
- **Filename**: `SPO_PermissionsMatrix_YYYYMMDD_HHMMSS.html`
- **Format**: Professional HTML report with responsive design
- **Features**:
  - Executive summary with key statistics
  - Color-coded permissions by type
  - Site-by-site detailed breakdown
  - Error and warning sections
  - Interactive navigation
  - Print-friendly layout

## Report Data Structure

Each permission entry includes:

| Field | Description |
|-------|-------------|
| `SiteUrl` | URL of the SharePoint site |
| `SiteTitle` | Display name of the site |
| `ObjectType` | Type of object (Site, Subsite, List, Library) |
| `ObjectTitle` | Display name of the object |
| `ObjectUrl` | Direct URL to the object |
| `PrincipalType` | Type of principal (User, Group) |
| `PrincipalName` | Display name of the user or group |
| `PrincipalLoginName` | Login name or identifier |
| `PermissionLevel` | SharePoint permission level |
| `PermissionType` | How permission was granted (Direct, SharePoint Group) |
| `IsInherited` | Whether permission is inherited |
| `SharingLinkType` | Type of sharing link (if applicable) |
| `ExpirationDate` | Expiration date for sharing links |
| `CreatedDate` | When the audit entry was created |

## Performance Considerations

### Throttling Guidelines
- **Small Tenants** (< 100 sites): ThrottleLimit = 5
- **Medium Tenants** (100-500 sites): ThrottleLimit = 3
- **Large Tenants** (500+ sites): ThrottleLimit = 2
- **Very Large Tenants** (1000+ sites): ThrottleLimit = 1

### Execution Time Estimates
- **Basic Site Audit**: ~30 seconds per site
- **With Subsites**: +50% additional time
- **With List Permissions**: +100% additional time
- **Full Comprehensive Audit**: ~2-3 minutes per site

### Memory Usage
- **Basic Audit**: ~50MB per 100 sites
- **Comprehensive Audit**: ~200MB per 100 sites
- **Large Tenancies**: Consider running in segments

## Error Handling

The script includes comprehensive error handling:

- **Connection Failures**: Automatic retry with detailed logging
- **Permission Denied**: Graceful handling with error reporting
- **API Throttling**: Built-in delays and retry logic
- **Module Issues**: Validation and installation guidance
- **Partial Failures**: Continue processing with error tracking

## Security Considerations

### Best Practices
- **Use Certificate Authentication** for production environments
- **Store Credentials Securely** using Azure Key Vault or similar
- **Regular Audits**: Schedule regular permission audits
- **Access Reviews**: Use reports for quarterly access reviews
- **Compliance**: Maintain audit trails for compliance requirements

### Data Privacy
- The script **does not** store passwords or sensitive data
- Reports contain user identities for legitimate business purposes
- Follow your organization's data retention policies
- Consider encrypting report storage locations

## Troubleshooting

### Common Issues

**Module Installation Issues**
```powershell
# Run as Administrator if needed
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Install-Module -Name PnP.PowerShell -Force -AllowClobber
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber

# If you get conflicts with older modules
Uninstall-Module -Name SharePointPnPPowerShellOnline -AllVersions -Force
Install-Module -Name PnP.PowerShell -Force
```

**Connection Issues - "Connect-SPOService not recognized"**
```powershell
# Manually import the module
Import-Module Microsoft.Online.SharePoint.PowerShell -Force

# Verify the module is loaded
Get-Module Microsoft.Online.SharePoint.PowerShell

# Check if cmdlet is available
Get-Command Connect-SPOService
```

**SharePoint Client Runtime Assembly Error**
```powershell
# Error: "Could not load type 'Microsoft.SharePoint.Client.SharePointOnlineCredentials'"
# Solution 1: Use PnP PowerShell only (recommended)
.\AuditSPO.PermissionsMatrix.ps1 -TenantName "contoso" -UsePnPOnly -GenerateHtmlReport

# Solution 2: Reinstall SharePoint Online Management Shell
Uninstall-Module Microsoft.Online.SharePoint.PowerShell -AllVersions
Install-Module Microsoft.Online.SharePoint.PowerShell -Force

# Solution 3: Clear module cache and reinstall
Remove-Item $env:USERPROFILE\Documents\PowerShell\Modules\Microsoft.Online.SharePoint.PowerShell -Recurse -Force
Install-Module Microsoft.Online.SharePoint.PowerShell -Force

# Solution 4: Download MSI installer from Microsoft
# Visit: https://www.microsoft.com/en-us/download/details.aspx?id=35588
```

**Connection Timeouts**
- Reduce `ThrottleLimit` parameter
- Check network connectivity
- Verify credentials and permissions

**Large Tenant Performance**
- Use `SiteFilter` to process in segments
- Enable `VerboseLogging` to monitor progress
- Consider running during off-peak hours

**Permission Denied Errors**
- Verify SharePoint Administrator role
- Check site collection administrator permissions
- Review Azure AD application permissions

## Integration

### PowerBI Integration
```powershell
# Generate CSV for PowerBI consumption
.\AuditSPO.PermissionsMatrix.ps1 -TenantName "contoso" -OutputPath "C:\PowerBI\Data"
```

### Scheduled Execution
```powershell
# Create scheduled task for weekly audits
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\Scripts\AuditSPO.PermissionsMatrix.ps1 -TenantName contoso -GenerateHtmlReport"
$Trigger = New-ScheduledTaskTrigger -Weekly -At 2AM -DaysOfWeek Sunday
Register-ScheduledTask -TaskName "SPO Permissions Audit" -Action $Action -Trigger $Trigger
```

### Log Analytics Integration
The script generates structured logs suitable for Azure Log Analytics ingestion for enterprise monitoring and alerting.

## Version History

- **v2.0** - Complete rewrite with enhanced features
  - Added HTML reporting
  - Implemented throttling mechanism
  - Enhanced error handling
  - Added multiple authentication methods
  - Improved performance optimization

- **v1.0** - Initial release
  - Basic CSV reporting
  - Simple permissions audit

## Support and Documentation

For detailed script execution guidance and troubleshooting:
- **Blog Post**: [SharePoint Online Permissions Matrix Audit Guide](https://educ4te.com/sharepoint/audit-sharepoint-online-permissions-matrix/)
- **Training Resources**: Available on EDUC4TE website
- **Community Support**: GitHub Issues and Discussions

## License

This script is provided as-is under the MIT License. See the repository license for full terms.

---

## About EDUC4TE

This script is part of the **EDUC4TE Microsoft 365 Solutions** collection. EDUC4TE provides:

- **Microsoft 365 Training** for comprehensive skills development
- **Consultancy Services** for enterprise Microsoft 365 implementations
- **Custom PowerShell Solutions** for automation and reporting
- **Security & Compliance Guidance** for regulatory requirements

**Learn More**: [Visit EDUC4TE](https://educ4te.com) | [Contact Us](https://educ4te.com/contact)

---

*Script authored by Alpesh Nakar - EDUC4TE*