# SharePoint Site Collections and Content Creation Script

## Overview

This PowerShell script automatically creates SharePoint Online site collections with subsites and populates them with Office documents for testing or demo purposes. The script creates 10 site collections by default, each containing 20 subsites, all populated with Word documents, Excel spreadsheets, and PDF files.

## Features

- ✅ **Modern Authentication Support** - Certificate-based authentication priority
- ✅ **MFA-Enabled Account Support** - Interactive authentication for MFA accounts
- ✅ **Automatic Module Installation** - PnP PowerShell module auto-installation
- ✅ **Progress Reporting** - Real-time progress tracking and error handling
- ✅ **Scheduler-Friendly** - Certificate authentication for automation
- ✅ **Bulk Document Creation** - Automated Office document generation
- ✅ **Comprehensive Reporting** - Detailed execution summary and error logs

## Prerequisites

### Required Permissions
- **SharePoint Administrator** or **Global Administrator** role
- **Application permissions** (if using certificate authentication):
  - `Sites.FullControl.All`
  - `Sites.Manage.All`

### PowerShell Requirements
- PowerShell 5.1 or PowerShell 7+
- Internet connectivity
- PnP PowerShell module (auto-installed if missing)

## Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `TenantUrl` | String | **Yes** | - | SharePoint tenant URL (e.g., `https://contoso.sharepoint.com`) |
| `ClientId` | String | No | - | Azure AD Application Client ID (for certificate auth) |
| `CertificateThumbprint` | String | No | - | Certificate thumbprint (for certificate auth) |
| `TenantId` | String | No | - | Azure AD Tenant ID (for certificate auth) |
| `UserName` | String | No | - | Username for basic authentication |
| `Password` | SecureString | No | - | Secure password for basic authentication |
| `SiteCollectionPrefix` | String | No | `TestSite` | Prefix for site collection names |
| `SiteCollectionCount` | Integer | No | `10` | Number of site collections to create |
| `SitesPerCollection` | Integer | No | `20` | Number of subsites per collection |
| `WordDocCount` | Integer | No | `30` | Word documents per site |
| `ExcelSheetCount` | Integer | No | `30` | Excel files per site |
| `PdfFileCount` | Integer | No | `50` | PDF files per site |

## Usage Examples

### 1. Interactive Authentication (Recommended for Testing)
```powershell
.\1.Create Site Content.ps1 -TenantUrl "https://contoso.sharepoint.com"
```

### 2. Interactive Authentication with Custom Settings
```powershell
.\1.Create Site Content.ps1 -TenantUrl "https://contoso.sharepoint.com" -SiteCollectionPrefix "Demo" -SiteCollectionCount 5 -SitesPerCollection 10
```

### 3. Certificate-Based Authentication (Recommended for Production)
```powershell
.\1.Create Site Content.ps1 -TenantUrl "https://contoso.sharepoint.com" -ClientId "12345678-1234-1234-1234-123456789012" -CertificateThumbprint "A1B2C3D4E5F6..." -TenantId "87654321-4321-4321-4321-210987654321"
```

### 4. Basic Authentication (Legacy)
```powershell
$SecurePassword = ConvertTo-SecureString "YourPassword" -AsPlainText -Force
.\1.Create Site Content.ps1 -TenantUrl "https://contoso.sharepoint.com" -UserName "admin@contoso.onmicrosoft.com" -Password $SecurePassword
```

### 5. Minimal Test Setup
```powershell
.\1.Create Site Content.ps1 -TenantUrl "https://contoso.sharepoint.com" -SiteCollectionCount 2 -SitesPerCollection 5 -WordDocCount 5 -ExcelSheetCount 5 -PdfFileCount 10
```

## Step-by-Step Execution Guide

### Step 1: Prepare Your Environment
1. Open PowerShell as Administrator
2. Navigate to the script directory:
   ```powershell
   cd "C:\Path\To\Script\Directory"
   ```

### Step 2: Set Execution Policy (if needed)
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Step 3: Run the Script
Choose one of the usage examples above based on your authentication method.

### Step 4: Monitor Progress
- The script will display real-time progress
- Watch for any error messages in red
- Progress bars will show completion status

### Step 5: Review Results
- Check the generated CSV report
- Review any error logs
- Verify created sites in SharePoint admin center

## Authentication Methods

### 1. Interactive Authentication (Default)
- **Best for**: Manual execution, testing, MFA-enabled accounts
- **Requires**: Browser authentication
- **Security**: Highest (supports MFA)

### 2. Certificate-Based Authentication
- **Best for**: Automation, scheduled tasks, production environments
- **Requires**: Azure AD app registration with certificate
- **Security**: High (certificate-based)

### 3. Basic Authentication
- **Best for**: Legacy scenarios (not recommended)
- **Requires**: Username and password
- **Security**: Lower (password-based)

## Azure AD App Registration (For Certificate Authentication)

### Create App Registration
1. Go to Azure portal → Azure Active Directory → App registrations
2. Click "New registration"
3. Name: "SharePoint Site Creator"
4. Redirect URI: Not required
5. Click "Register"

### Configure API Permissions
1. Go to API permissions
2. Add permissions:
   - **SharePoint** → Application permissions → `Sites.FullControl.All`
   - **Microsoft Graph** → Application permissions → `Sites.FullControl.All`
3. Grant admin consent

### Upload Certificate
1. Go to Certificates & secrets
2. Upload your certificate (.cer file)
3. Note the thumbprint

## Output and Reporting

### Generated Files
- **CSV Report**: `SharePoint_Provisioning_Report_[timestamp].csv`
- **Console Output**: Real-time progress and status messages
- **Error Logs**: Captured in the global configuration

### Report Contents
- Site collection details (URL, title, creation time)
- Subsite information
- Document creation summary
- Error details (if any)

## Troubleshooting

### Common Issues

#### 1. Module Installation Fails
```powershell
# Manual installation
Install-Module PnP.PowerShell -Scope CurrentUser -Force
```

#### 2. Authentication Issues
```powershell
# Check if you're already connected
Get-PnPConnection

# Disconnect and retry
Disconnect-PnPOnline
```

#### 3. Permission Errors
- Verify you have SharePoint Administrator rights
- Check Azure AD app permissions
- Ensure certificate is properly configured

#### 4. Site Already Exists
- The script will skip existing sites
- Check the error log for details
- Consider using a different prefix

### Debug Mode
```powershell
# Run with verbose output
.\1.Create Site Content.ps1 -TenantUrl "https://contoso.sharepoint.com" -Verbose
```

## Performance Considerations

### Resource Usage
- **Memory**: Moderate (document templates stored in memory)
- **Network**: High (multiple SharePoint API calls)
- **Time**: Varies by site count (estimate 2-5 minutes per site collection)

### Optimization Tips
1. **Reduce Document Counts** for faster execution
2. **Use Certificate Authentication** for better performance
3. **Run During Off-Peak Hours** to avoid throttling
4. **Monitor Network Bandwidth** during execution

## Security Best Practices

### 1. Authentication
- ✅ Use certificate-based authentication for automation
- ✅ Enable MFA for interactive sessions
- ❌ Avoid storing passwords in scripts

### 2. Permissions
- ✅ Use principle of least privilege
- ✅ Regularly review app registrations
- ✅ Remove unused certificates

### 3. Monitoring
- ✅ Review execution logs
- ✅ Monitor created sites
- ✅ Set up alerts for failures

## Cleanup

### Remove Created Sites
```powershell
# Connect to SharePoint admin center
Connect-PnPOnline -Url "https://contoso-admin.sharepoint.com" -Interactive

# List and remove sites (be careful!)
Get-PnPTenantSite | Where-Object {$_.Title -like "TestSite*"} | Remove-PnPTenantSite -Force
```

### Disconnect Sessions
```powershell
Disconnect-PnPOnline
```

## Support and Resources

### Microsoft Documentation
- [PnP PowerShell Documentation](https://pnp.github.io/powershell/)
- [SharePoint Online Management Shell](https://docs.microsoft.com/en-us/powershell/sharepoint/)
- [Azure AD App Registration](https://docs.microsoft.com/en-us/azure/active-directory/develop/)

### Related Scripts in This Repository
- [`Install-Microsoft365Modules.ps1`](../../../00.Core/Install-Microsoft365Modules.ps1) - Install required modules
- [`Verify-M365Modules.ps1`](../../../00.Core/Verify-M365Modules.ps1) - Verify module installation
- [`Microsoft365-Commands-ReadyToUse.md`](../../../00.Core/Microsoft365-Commands-ReadyToUse.md) - Common commands reference

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 3.0 | Current | Modern authentication, certificate support, enhanced error handling |
| 2.0 | Previous | Added document creation capabilities |
| 1.0 | Initial | Basic site collection creation |

---

## ⚠️ Important Notes

1. **Test Environment**: Always test in a development environment first
2. **Backup**: Ensure you have proper backups before running in production
3. **Quotas**: Be aware of SharePoint storage quotas
4. **Throttling**: Microsoft may throttle requests during bulk operations
5. **Licensing**: Ensure users have appropriate SharePoint licenses

## 📞 Support

If you encounter issues:
1. Check the error logs generated by the script
2. Review the troubleshooting section above
3. Verify permissions and authentication setup
4. Consult Microsoft documentation for specific error codes

---

*This script is part of the AdminDroid Community PowerShell Scripts collection for Microsoft 365 administration.*