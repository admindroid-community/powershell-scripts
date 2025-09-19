# SharePoint Online Permissions Matrix Audit Tool v2.0 - Interactive Usage Examples
# Author: Alpesh Nakar | EDUC4TE | educ4te.com
# Updated: September 19, 2025

<#
============================================================================================
INTERACTIVE AUTHENTICATION EXAMPLES - PREFERRED METHOD
============================================================================================
Interactive authentication is now the highest priority method and supports:
- Multi-Factor Authentication (MFA)
- Modern Authentication
- Conditional Access policies
- Browser-based authentication flow
============================================================================================
#>

# Example 1: Basic Interactive Authentication (Recommended)
# This will prompt for interactive login with MFA support
& ".\AuditSPO.PermissionsMatrix_v2.ps1" `
    -TenantName "contoso" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -GenerateHtmlReport `
    -VerboseLogging

# Example 2: Interactive Authentication with All Features
# Comprehensive audit with modern authentication
& ".\AuditSPO.PermissionsMatrix_v2.ps1" `
    -TenantName "contoso" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -ThrottleLimit 3 `
    -IncludeSubsites `
    -IncludeListPermissions `
    -IncludeSharingLinks `
    -GenerateHtmlReport `
    -OutputPath "C:\SharePointAudit" `
    -VerboseLogging

# Example 3: Quick Site Filter Audit
# Focus on specific sites with interactive auth
& ".\AuditSPO.PermissionsMatrix_v2.ps1" `
    -TenantName "contoso" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -SiteFilter @("finance", "hr", "executive") `
    -GenerateHtmlReport

# Example 4: Large Tenant Conservative Approach
# Lower throttle for large environments with interactive auth
& ".\AuditSPO.PermissionsMatrix_v2.ps1" `
    -TenantName "contoso" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -ThrottleLimit 1 `
    -GenerateHtmlReport `
    -OutputPath "C:\LargeTenantAudit"

<#
============================================================================================
FALLBACK AUTHENTICATION METHODS
============================================================================================
These methods are available as fallbacks but interactive is preferred
============================================================================================
#>

# Example 5: Certificate-Based Authentication (For Automation)
& ".\AuditSPO.PermissionsMatrix_v2.ps1" `
    -TenantName "contoso" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -CertificateThumbprint "1234567890ABCDEF1234567890ABCDEF12345678" `
    -GenerateHtmlReport `
    -ThrottleLimit 2

# Example 6: Credential-Based Authentication (Legacy/Service Accounts)
# Note: Password is now SecureString for better security
$SecurePassword = ConvertTo-SecureString "YourPassword" -AsPlainText -Force
& ".\AuditSPO.PermissionsMatrix_v2.ps1" `
    -TenantName "contoso" `
    -UserName "admin@contoso.onmicrosoft.com" `
    -Password $SecurePassword `
    -GenerateHtmlReport

<#
============================================================================================
AUTHENTICATION PRIORITY ORDER (NEW IN V2.0)
============================================================================================
1. Interactive Authentication (Highest Priority)
   - Best for manual audits
   - Supports MFA and Conditional Access
   - Modern browser-based authentication

2. Certificate-Based Authentication
   - Best for automation and scheduled tasks
   - No password exposure
   - Enterprise-grade security

3. Credential-Based Authentication (Fallback)
   - For legacy environments
   - Service accounts without certificate infrastructure
   - Now uses SecureString for better security
============================================================================================
#>

Write-Host "🚀 SharePoint Online Permissions Matrix Audit Tool v2.0" -ForegroundColor Green
Write-Host "✨ Interactive authentication is now the preferred method!" -ForegroundColor Cyan
Write-Host "📖 Choose an example above and customize for your environment" -ForegroundColor Yellow
Write-Host "🌐 For more information: https://educ4te.com" -ForegroundColor Blue