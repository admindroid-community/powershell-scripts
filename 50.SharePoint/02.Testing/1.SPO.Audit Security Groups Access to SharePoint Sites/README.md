## Audit Microsoft Security Group Access to SharePoint Sites Using PowerShell

This PowerShell script helps administrators audit Microsoft Security Group permissions across all SharePoint Online sites in a tenant. The script provides comprehensive visibility into where security groups have access and what permission levels they possess.

***Sample Output:***

The script exports an output CSV file with the following information:
- Site URL and Title
- Security Group Name and Type  
- Permission Levels (Full Control, Edit, Read, etc.)
- Permission Type (Direct or Inherited)
- Creation Date of the audit

## Script Highlights

- ✅ **Modern Authentication**: Supports certificate-based authentication (CBA) and MFA-enabled accounts
- ✅ **Automatic Module Installation**: Installs PnP.PowerShell module upon confirmation
- ✅ **Comprehensive Analysis**: Reviews all SharePoint sites for the specified security group
- ✅ **Detailed Permissions**: Shows exact permission levels (Full Control, Edit, Read, etc.)
- ✅ **Subsite Support**: Optional analysis of subsites and nested permissions
- ✅ **System Site Filtering**: Excludes system and administrative sites by default
- ✅ **Progress Tracking**: Real-time progress updates during site analysis
- ✅ **Error Handling**: Comprehensive error logging and warning system
- ✅ **CSV Export**: Detailed results exported to timestamped CSV file
- ✅ **Scheduler Friendly**: Perfect for automated auditing scenarios
- ✅ **Flexible Group Matching**: Case-insensitive matching for security group names

## Prerequisites

1. **PowerShell Modules Required:**
   - `PnP.PowerShell` (version 1.12.0 or later)

2. **Permissions Required:**
   - SharePoint Online Administrator or Global Administrator
   - Azure AD App Registration with appropriate permissions (for certificate authentication)

3. **Authentication Options:**
   - Interactive authentication (supports MFA) - **Requires ClientId**
   - Certificate-based authentication for automation
   - Basic authentication (legacy, not recommended)

## Parameters

| Parameter | Type | Description | Required |
|-----------|------|-------------|----------|
| `SecurityGroupName` | String | Name of the Microsoft Security Group to audit | ✅ Yes |
| `TenantName` | String | SharePoint tenant name (e.g., contoso) | No |
| `ClientId` | String | Azure AD App Client ID for certificate auth | No |
| `CertificateThumbprint` | String | Certificate thumbprint for authentication | No |
| `UserName` | String | Username for basic authentication | No |
| `Password` | String | Password for basic authentication | No |
| `IncludeSubsites` | Switch | Include analysis of subsites | No |
| `IncludeSystemSites` | Switch | Include system and administrative sites | No |

## Usage Examples

### Example 1: Interactive Authentication (Recommended)
```powershell
.\AuditSecurityGroupSharePointAccess.ps1 -SecurityGroupName "Finance Team" -TenantName "contoso"
```
**Note:** Interactive authentication will prompt for a ClientId if not provided as a parameter.

### Example 2: Certificate-Based Authentication (Automation)
```powershell
.\AuditSecurityGroupSharePointAccess.ps1 -SecurityGroupName "HR Department" -TenantName "contoso" -ClientId "12345678-1234-1234-1234-123456789012" -CertificateThumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12"
```

### Example 3: With ClientId for Interactive Authentication
```powershell
.\AuditSecurityGroupSharePointAccess.ps1 -SecurityGroupName "Marketing Team" -TenantName "contoso" -ClientId "afe1b358-534b-4c96-abb9-ecea5d5f2e5d"
```

### Example 4: Include Subsites Analysis
```powershell
.\AuditSecurityGroupSharePointAccess.ps1 -SecurityGroupName "IT Team" -IncludeSubsites -TenantName "contoso" -ClientId "afe1b358-534b-4c96-abb9-ecea5d5f2e5d"
```

### Example 5: Include System Sites
```powershell
.\AuditSecurityGroupSharePointAccess.ps1 -SecurityGroupName "IT Administrators" -IncludeSystemSites -TenantName "contoso"
```

## Sample Output

The script generates a CSV report with the following columns:

| Column | Description |
|--------|-------------|
| SiteUrl | Full URL of the SharePoint site |
| SiteTitle | Display name of the SharePoint site |
| GroupName | Name of the security group |
| GroupType | Type of group (SecurityGroup, SharePointGroup) |
| LoginName | Internal login name/identifier |
| PermissionLevels | Assigned permission levels (comma-separated) |
| PermissionType | Direct or Inherited permissions |
| CreatedDate | Timestamp when the audit was performed |

## Common Permission Levels

- **Full Control**: Complete access to site and content
- **Design**: Modify site structure and appearance  
- **Edit**: Create, edit, and delete content
- **Contribute**: Add and edit content
- **Read**: View content only
- **View Only**: View content without downloading

## Troubleshooting

**Issue: "Please specify a valid client id for an Entra ID App Registration"**
- Interactive authentication now requires a ClientId parameter
- Provide a valid Azure AD App Registration ClientId
- Example: `-ClientId "afe1b358-534b-4c96-abb9-ecea5d5f2e5d"`

**Issue: "Security group not found"**
- Verify the exact spelling of the security group name (matching is case-insensitive)
- The script searches for groups like "Everyone", "Everyone except external users", and custom security groups
- Try searching for partial names or common system groups first

**Issue: "Insufficient permissions to read site"**
- Ensure you have SharePoint Administrator rights
- Some sites may have unique permissions that prevent access
- Consider running as Global Administrator
- Check the error log for specific sites that couldn't be accessed

**Issue: "Module installation fails"**
- Run PowerShell as Administrator for module installation
- Use `-Scope CurrentUser` if you don't have admin rights
- Check your execution policy: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

**Issue: "Attempted to perform an unauthorized operation"**
- This typically occurs on specific sites with restricted access
- The script will continue processing other sites and log this as an error
- Review the error summary at the end of execution

## Best Practices

1. **Regular Auditing**: Schedule monthly audits to track permission changes
2. **Least Privilege**: Review results to ensure groups have minimal necessary permissions
3. **Documentation**: Keep audit reports for compliance and security reviews
4. **Automation**: Use certificate-based authentication for scheduled runs
5. **Filtering**: Use system site exclusion to focus on business-relevant sites
6. **Testing**: Start with known security groups like "Everyone" to verify script functionality
7. **ClientId Management**: Store your Azure AD App ClientId securely for reuse
8. **Error Review**: Always check the error summary for sites that couldn't be accessed

## Security Considerations

- Store certificates securely when using certificate-based authentication
- Rotate authentication certificates regularly
- Review audit logs for any unauthorized permission changes
- Consider using Azure AD Privileged Identity Management for elevated access
- Protect your Azure AD App ClientId - while not a secret, it should be managed securely
- Regularly review and rotate authentication credentials

## Performance Notes

- The script connects to each individual SharePoint site for detailed permission analysis
- Execution time depends on the number of sites in your tenant (typically 1-3 seconds per site)
- For large tenants, consider running during off-hours
- Certificate-based authentication is faster than interactive authentication for bulk operations

---

*This script is part of the educ4te.com PowerShell script collection for Microsoft 365 administration and auditing.*
