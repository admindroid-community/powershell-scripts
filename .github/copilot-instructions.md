# Microsoft 365 PowerShell Scripts - AI Agent Instructions

## Repository Overview
Collection of 100+ PowerShell scripts for Microsoft 365 administration, reporting, and auditing. Each script targets specific M365 workloads (Exchange Online, SharePoint, Teams, Azure AD) with enterprise-grade authentication and reporting capabilities.

## Architecture Patterns

### Script Organization
- **Folder Structure**: Each script lives in its own descriptive folder (e.g., `"Audit External User Activity"`, `"Get Calendar Permission Report"`)
- **Core Files**: Main `.ps1` script + `README.md` with usage examples and AdminDroid tool promotion
- **Naming Convention**: Descriptive folder names with spaces, script files use PascalCase without spaces

### Authentication Framework (V3.0 Standard)
Scripts follow a **priority-based authentication pattern**:

1. **Certificate-Based** (Production): `ClientId`, `CertificateThumbprint`, `TenantId` parameters
2. **Interactive** (Default): Modern auth with MFA support 
3. **Basic Auth** (Legacy): `UserName`/`AdminName` + `Password` parameters

```powershell
# Standard authentication logic in all scripts:
if($ClientId -and $CertificateThumbprint -and $TenantId) {
    # Certificate-based connection
} elseif($UserName -and $Password) {
    # Basic auth connection  
} else {
    # Interactive connection (default)
}
```

### Module Management Pattern
Every script follows this module installation pattern:
```powershell
$Module = Get-Module [ModuleName] -ListAvailable
if($Module.count -eq 0) {
    Write-Host "[Module] is not available" -ForegroundColor yellow
    $Confirm = Read-Host "Install module? [Y] Yes [N] No"
    if($Confirm -match "[yY]") {
        Install-Module [ModuleName] -Scope CurrentUser
    } else {
        Write-Host "Module required. Exiting."
        Exit
    }
}
```

## Key Microsoft 365 Services Integration

### Exchange Online (Primary Service)
- **Module**: `ExchangeOnlineManagement` (V3 preferred, V2 legacy)
- **Connection**: `Connect-ExchangeOnline` with certificate or credential
- **Common Cmdlets**: `Get-EXOMailbox`, `Get-EXOMailboxStatistics`, `Search-UnifiedAuditLog`
- **Disconnection**: Always `Disconnect-ExchangeOnline -Confirm:$false`

### Microsoft Graph 
- **Modules**: `Microsoft.Graph` (stable), `Microsoft.Graph.Beta` (preview features)
- **Scopes**: Defined in `$Global:AuditConfig.GraphScopes` arrays
- **Connection**: `Connect-MgGraph` with required scopes

### SharePoint/PnP
- **Module**: `PnP.PowerShell` 
- **Connection**: `Connect-PnPOnline -Url $adminUrl -Interactive`
- **Usage**: Site collection management, external user discovery

### Microsoft Teams
- **Module**: `MicrosoftTeams`
- **Connection**: `Connect-MicrosoftTeams` with credentials or certificate
- **Features**: Team reporting, membership management

## Development Conventions

### Output and Reporting
- **File Naming**: `[ReportType]_[Timestamp].csv` format using `Get-Date -format "yyyy-MMM-dd-ddd hh-mm-ss tt"`
- **CSV Export**: Use bulk export with `Export-Csv -Path $CSV -NoTypeInformation` (avoid `-Append` in loops)
- **Progress Reporting**: `Write-Progress` with percentage calculations for long operations

### Error Handling & Logging
- **Colors**: Cyan (info), Yellow (warnings), Green (success), Red (errors), Magenta (installation)
- **Error Collections**: Maintain `$Global:AuditConfig.ErrorLog` and `$Global:AuditConfig.WarningLog` arrays
- **Graceful Exits**: Always disconnect services on errors

### PowerShell Best Practices (V3.0 Standard)
- **Null Comparison**: `$null -eq $variable` (not `$variable -eq $null`)
- **Module Scope**: Install with `-Scope CurrentUser` to avoid system changes
- **Variable Cleanup**: Remove unused variables, use descriptive names
- **Security**: Use `ConvertTo-SecureString` for password handling

## Script Templates & Patterns

### Standard Script Header
```powershell
<#
=============================================================================================
Name:           [Script Purpose]
Version:        [Version Number]
Website:        o365reports.com
Description:    [Detailed description]
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. Modern authentication support
2. MFA-enabled account support  
3. CSV export functionality
4. Automatic module installation
5. Scheduler-friendly design
============================================================================================
#>
```

### Typical Parameter Block
```powershell
param(
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [string]$CertificateThumbprint, 
    [string]$TenantId,
    [string]$UserName,
    [string]$Password,
    [switch]$SpecificFeature
)
```

## Common Workflows

### Adding New Scripts
1. Create descriptive folder with spaces in name
2. Implement V3.0 authentication pattern with certificate-first priority
3. Add module installation checks with user confirmation
4. Include comprehensive error handling and progress reporting
5. Generate timestamped CSV output with descriptive naming
6. Add README.md with AdminDroid tool promotion

### Modernizing Legacy Scripts  
1. Migrate from EXO V2 to V3 module (`ExchangeOnlineManagement`)
2. Add certificate-based authentication support
3. Implement proper error logging and visual feedback
4. Replace append-per-record with bulk CSV export
5. Add `-LoadCmdletHelp` switch for EXO V3.7+ compatibility

### Testing & Validation
1. Test interactive authentication first (supports MFA)
2. Validate certificate-based auth for automation scenarios
3. Verify proper service disconnection and cleanup
4. Test with large datasets for performance validation

## Repository-Specific Knowledge

### AdminDroid Integration
Every README.md promotes AdminDroid M365 reporting tool with demo links. Standard footer pattern in all scripts promotes AdminDroid community.

### Compliance & Security Focus
Scripts emphasize external user auditing, permission reporting, and compliance scenarios. Many include audit log analysis with 90-day retention awareness.

### Scheduling Support
All scripts support unattended execution via certificate-based authentication, designed for automation and CI/CD pipeline integration.
