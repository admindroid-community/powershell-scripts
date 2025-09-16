# Enhanced Security Group SharePoint Access Audit Script v2.0

This PowerShell script provides comprehensive auditing of Microsoft Security Group access across SharePoint Online sites, matching enterprise audit report formats.

## Features

✅ **Enhanced Audit Matrix**: Generates detailed permission matrix with granular permission levels (Full Control, Design, Contribute, Edit, Read, etc.)

✅ **Comprehensive Coverage**: Audits site collections, subsites, document libraries, and lists

✅ **Inheritance Analysis**: Shows whether permissions are directly assigned or inherited

✅ **Security Group Mapping**: Identifies which SharePoint groups contain the security groups

✅ **Enterprise Format**: Output matches AdminDroid-style audit reports with detailed columns

✅ **Flexible Targeting**: Can audit specific security groups or all security groups

✅ **Modern Authentication**: Supports certificate-based authentication for automation

## Script Highlights

- Uses modern authentication to connect to SharePoint Online
- Supports certificate-based authentication (CBA) for automation scenarios
- Supports MFA-enabled account authentication
- Automatically installs required PnP.PowerShell module upon confirmation
- Audits security group permissions across all SharePoint sites, lists, and libraries
- Exports detailed permission matrix report matching enterprise audit formats
- Identifies direct and inherited permissions with inheritance details
- Shows granular permission levels with "Given through" information
- Includes site collections, subsites, lists, and document libraries analysis
- Scheduler-friendly design for automated auditing
- Comprehensive error handling and progress reporting

## Parameters

| Parameter | Type | Description | Required |
|-----------|------|-------------|----------|
| SecurityGroupName | String | Name of specific security group to audit (leave empty for all groups) | No |
| TenantName | String | SharePoint tenant name (e.g., contoso) | No |
| ClientId | String | Azure App Client ID for certificate authentication | No |
| CertificateThumbprint | String | Certificate thumbprint for authentication | No |
| UserName | String | Admin username for basic authentication | No |
| Password | String | Admin password for basic authentication | No |
| IncludeSubsites | Switch | Include subsites in the audit | No |
| IncludeSystemSites | Switch | Include system sites (search, admin, etc.) | No |
| IncludeListsAndLibraries | Switch | Include document libraries and lists analysis | No |

## Output Format

The script generates a CSV file with the following columns matching enterprise audit standards:

- **Type**: Resource type (Microsoft 365 tenant, Site collection, Document library, List)
- **Name**: Display name of the resource
- **URL**: Full URL to the resource
- **Item path**: Path within the resource (if applicable)
- **Inheritance**: Whether permissions are Custom or Inherited
- **Details**: Additional details about the permission
- **User/group**: Name of the security group
- **Principal type**: Type of principal (Security group)
- **Account name**: Technical account name/claim
- **Given through**: SharePoint group through which access is granted
- **Full Control**: X if has Full Control permission
- **Design**: X if has Design permission
- **Contribute**: X if has Contribute permission
- **Edit**: X if has Edit permission
- **Read**: X if has Read permission
- **Restricted View**: X if has Restricted View permission
- **Web-Only Limited Access**: X if has Limited Access permission

## Usage Examples

### Example 1: Interactive Authentication (Audit All Security Groups)
```powershell
.\AuditSecurityGroupSharePointAccess_v2.ps1 -TenantName "contoso" -IncludeListsAndLibraries
```

### Example 2: Audit Specific Security Group
```powershell
.\AuditSecurityGroupSharePointAccess_v2.ps1 -SecurityGroupName "Leadership Members" -TenantName "contoso" -IncludeListsAndLibraries -IncludeSubsites
```

### Example 3: Certificate-Based Authentication (Automation)
```powershell
.\AuditSecurityGroupSharePointAccess_v2.ps1 -ClientId "your-client-id" -CertificateThumbprint "certificate-thumbprint" -TenantName "contoso" -IncludeListsAndLibraries
```

### Example 4: Basic Authentication
```powershell
.\AuditSecurityGroupSharePointAccess_v2.ps1 -UserName "admin@contoso.com" -Password "SecurePassword" -TenantName "contoso" -SecurityGroupName "IT Security"
```

### Example 5: Comprehensive Audit with All Options
```powershell
.\AuditSecurityGroupSharePointAccess_v2.ps1 -TenantName "contoso" -IncludeSubsites -IncludeListsAndLibraries -IncludeSystemSites
```

## Prerequisites

- **PowerShell 5.1** or later
- **PnP.PowerShell** module (automatically installed if missing)
- **SharePoint Online Administrator** or **Global Administrator** permissions
- For certificate authentication: **Registered Azure App** with appropriate permissions

## Required Permissions

The script requires the following SharePoint permissions:
- Sites.Read.All (to read site collections)
- Sites.FullControl.All (to read permissions on all sites)
- User.Read.All (to resolve security group information)

## Output Files

- **CSV Report**: `SecurityGroupSharePointAccessMatrix_[timestamp].csv`
- **Format**: Enterprise-ready audit report with permission matrix
- **Location**: Current script directory

## Automation Support

The script supports unattended execution using certificate-based authentication:

1. Register an Azure App in your tenant
2. Upload a certificate to the Azure App
3. Grant required SharePoint permissions
4. Use the ClientId and CertificateThumbprint parameters

## Advanced Features

### Permission Matrix Analysis
- Maps SharePoint permission levels to standard enterprise audit format
- Shows exact permission inheritance chain
- Identifies direct vs inherited permissions

### Security Group Mapping
- Resolves security groups to SharePoint groups
- Shows complete permission delegation chain
- Identifies nested group memberships

### Comprehensive Coverage
- Site collections and subsites
- Document libraries and lists
- Custom permission levels
- System and hidden sites (optional)

## Troubleshooting

### Common Issues

1. **"Access Denied" errors**: Ensure you have SharePoint Administrator permissions
2. **"PnP PowerShell not found"**: Script will automatically prompt to install the module
3. **"No permissions found"**: Check group name spelling and ensure the group has SharePoint access
4. **Certificate authentication fails**: Verify certificate is installed and app permissions are granted

### Performance Optimization

- Use `-IncludeListsAndLibraries` only when detailed library analysis is needed
- Exclude system sites unless specifically required for compliance
- Use certificate authentication for faster automated runs

## Version History

- **v2.0**: Enhanced output format matching enterprise audit standards, added permission matrix, improved inheritance analysis
- **v1.0**: Basic security group permission auditing

## Support

For additional support and advanced M365 auditing tools, visit:
- **Website**: [educ4te.com](https://educ4te.com)


*This script is part of the comprehensive M365 PowerShell Scripts collection for enterprise administration and compliance reporting.*
