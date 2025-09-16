# SharePoint Online Tenant Rename Script

## Overview
This PowerShell script provides a comprehensive solution for renaming SharePoint Online tenants with enterprise-grade features including scheduled execution, comprehensive validation, and detailed logging.

## Features
- **Modern Authentication**: Certificate-based authentication with fallback to interactive auth
- **MFA Support**: Compatible with Multi-Factor Authentication enabled accounts
- **Scheduled Execution**: Configurable delay for planned maintenance windows
- **Validation Mode**: Pre-check tenant eligibility without performing rename
- **Comprehensive Logging**: Detailed logging with timestamps and error tracking
- **PowerShell Compatibility**: Automatic detection and handling of PowerShell version requirements

## Prerequisites
- **SharePoint Online Administrator** or **Global Administrator** permissions
- **PowerShell 5.1+** (PowerShell 7.4+ recommended for full feature support)
- **Microsoft.Online.SharePoint.PowerShell** module
- **PnP.PowerShell** module (optional, requires PowerShell 7.4+)

## Parameters

### Authentication Parameters
- `ClientId`: Azure AD Application Client ID (for certificate-based auth)
- `CertificateThumbprint`: Certificate thumbprint (for certificate-based auth)
- `TenantId`: Azure AD Tenant ID (for certificate-based auth)
- `UserName`: Username for credential-based authentication
- `SecurePassword`: Secure string password for credential-based authentication

### Tenant Configuration
- `CurrentTenantName`: Current SharePoint tenant name (default: "n4k4r")
- `NewTenantName`: New SharePoint tenant name (default: "educ4teaustralia")

### Execution Options
- `DelayHours`: Hours to wait before executing rename (0-168, default: 25)
- `Force`: Skip confirmation prompts
- `ValidationOnly`: Perform eligibility checks only without renaming

## Usage Examples

### Interactive Authentication (Default)
```powershell
.\spo.tenant.rename.ps1 -CurrentTenantName "oldtenant" -NewTenantName "newtenant"
```

### Certificate-based Authentication (Recommended for Production)
```powershell
.\spo.tenant.rename.ps1 -ClientId "your-client-id" -CertificateThumbprint "cert-thumbprint" -TenantId "tenant-id" -CurrentTenantName "oldtenant" -NewTenantName "newtenant"
```

### Validation Only Mode
```powershell
.\spo.tenant.rename.ps1 -CurrentTenantName "oldtenant" -NewTenantName "newtenant" -ValidationOnly
```

### Immediate Execution with Force
```powershell
.\spo.tenant.rename.ps1 -CurrentTenantName "oldtenant" -NewTenantName "newtenant" -DelayHours 0 -Force
```

## Security Best Practices

### Authentication Priority
1. **Certificate-based** (Production environments)
2. **Interactive** (Development/testing)
3. **Credential-based** (Legacy scenarios)

### Certificate Setup
For production deployments, configure certificate-based authentication:
1. Create Azure AD application registration
2. Upload certificate to Azure AD application
3. Grant SharePoint administrator permissions
4. Use certificate thumbprint for authentication

### Password Security
- Use `SecureString` for password parameters
- Avoid plain text passwords in scripts
- Consider using Windows Credential Manager integration

## Validation Checks
The script performs comprehensive pre-rename validation:
- **Tenant Verification**: Confirms current tenant name matches
- **Active Operations**: Checks for in-progress SharePoint operations
- **Site Collection Count**: Warns if high number of sites (>10,000)
- **PowerShell Compatibility**: Validates module requirements

## Monitoring and Logging
- **Real-time Progress**: Progress bars for long-running operations
- **Detailed Logging**: Timestamped log files with error categorization
- **Operation Tracking**: Monitors rename operation status every 5 minutes
- **Summary Reports**: Execution summary with error/warning counts

## Error Handling
- **Graceful Failures**: Proper cleanup on script interruption
- **Service Disconnection**: Automatic disconnection from SharePoint services
- **Error Categorization**: Separate tracking of errors and warnings
- **Retry Logic**: Built-in retry for transient connection issues

## PowerShell Version Compatibility
- **PowerShell 5.1**: Basic functionality with SPO module only
- **PowerShell 7.4+**: Full functionality including PnP.PowerShell integration
- **Automatic Detection**: Script detects version and adjusts accordingly

## Important Considerations

### Pre-Rename Checklist
1. **Backup Critical Data**: Ensure recent backups of critical SharePoint content
2. **Communicate Changes**: Notify users of upcoming domain changes
3. **Update Bookmarks**: Users will need to update saved links
4. **Review Dependencies**: Check external integrations using SharePoint URLs
5. **Plan Maintenance Window**: Rename process can take several hours

### Post-Rename Tasks
1. **Verify Functionality**: Test critical SharePoint functionality
2. **Update Documentation**: Update any documentation with new URLs
3. **Monitor Operations**: Watch for any issues in the hours following rename
4. **Update Configurations**: Update any systems referencing the old domain

## Troubleshooting

### Common Issues
- **Module Import Failures**: Ensure latest module versions are installed
- **Authentication Errors**: Verify credentials and permissions
- **Operation Timeouts**: Large tenants may require extended processing time
- **PowerShell Version**: Use PowerShell 7.4+ for full functionality

### Log File Analysis
Check the generated log file for detailed error information:
```
SPOTenantRename_YYYY-MMM-DD-DDD_HH-MM-SS_TT.log
```

## Support and Resources
- Microsoft SharePoint Online documentation
- PowerShell module documentation for troubleshooting
- Azure AD application registration guides for certificate setup

---
*This script follows Microsoft 365 security best practices and PowerShell coding standards.*
