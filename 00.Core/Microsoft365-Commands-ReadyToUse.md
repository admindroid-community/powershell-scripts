# Microsoft 365 PowerShell Commands - Ready to Use

This file contains ready-to-use PowerShell commands for Microsoft 365 administration. Replace the variables in `{BRACKETS}` with your actual values.

## 📋 Variables to Replace

Before using these commands, replace these variables with your actual values:

| Variable | Description | Example |
|----------|-------------|---------|
| `{TENANT_NAME}` | Your Microsoft 365 tenant name | `contoso` |
| `{TENANT_DOMAIN}` | Your primary domain | `contoso.onmicrosoft.com` |
| `{USER_EMAIL}` | User email address | `john.doe@contoso.com` |
| `{GROUP_NAME}` | Group display name | `Sales Team` |
| `{SITE_URL}` | SharePoint site URL | `https://contoso.sharepoint.com/sites/sales` |
| `{TEAM_NAME}` | Teams team name | `Marketing Department` |
| `{MAILBOX_NAME}` | Mailbox name or email | `shared@contoso.com` |
| `{ENVIRONMENT_NAME}` | Power Platform environment | `Development` |
| `{DEVICE_NAME}` | Device name | `DESKTOP-ABC123` |

---

## 🔐 Microsoft Graph (Entra ID & Microsoft 365)

### Initial Connection
```powershell
# Connect with basic permissions
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All"

# Connect with administrative permissions
Connect-MgGraph -Scopes @(
    "User.ReadWrite.All",
    "Group.ReadWrite.All", 
    "Directory.ReadWrite.All",
    "Application.ReadWrite.All",
    "RoleManagement.ReadWrite.Directory"
)

# Check connection status
Get-MgContext

# Disconnect when finished
Disconnect-MgGraph
```

### User Management
```powershell
# List all users
Get-MgUser -All | Select-Object DisplayName, UserPrincipalName, Id

# Get specific user
Get-MgUser -Filter "UserPrincipalName eq '{USER_EMAIL}'"

# Get user by display name
Get-MgUser -Filter "DisplayName eq 'John Doe'"

# Create new user
$passwordProfile = @{
    Password = "TempPassword123!"
    ForceChangePasswordNextSignIn = $true
}

New-MgUser -DisplayName "{DISPLAY_NAME}" `
           -UserPrincipalName "{USER_EMAIL}" `
           -MailNickname "{MAIL_NICKNAME}" `
           -PasswordProfile $passwordProfile `
           -AccountEnabled:$true

# Update user
Update-MgUser -UserId "{USER_EMAIL}" -JobTitle "Senior Developer" -Department "IT"

# Disable user
Update-MgUser -UserId "{USER_EMAIL}" -AccountEnabled:$false

# Delete user
Remove-MgUser -UserId "{USER_EMAIL}"

# Get user's group memberships
Get-MgUserMemberOf -UserId "{USER_EMAIL}"

# Get users created in last 30 days
$since = (Get-Date).AddDays(-30).ToString('yyyy-MM-ddTHH:mm:ssZ')
Get-MgUser -Filter "createdDateTime ge $since" | Select-Object DisplayName, CreatedDateTime
```

### Group Management
```powershell
# List all groups
Get-MgGroup -All | Select-Object DisplayName, Id, GroupTypes

# Get specific group
Get-MgGroup -Filter "DisplayName eq '{GROUP_NAME}'"

# Create new Microsoft 365 group
New-MgGroup -DisplayName "{GROUP_NAME}" `
            -MailNickname "{GROUP_NICKNAME}" `
            -MailEnabled:$true `
            -SecurityEnabled:$false `
            -GroupTypes @("Unified")

# Create security group
New-MgGroup -DisplayName "{GROUP_NAME}" `
            -MailEnabled:$false `
            -SecurityEnabled:$true

# Add user to group
$groupId = (Get-MgGroup -Filter "DisplayName eq '{GROUP_NAME}'").Id
$userId = (Get-MgUser -Filter "UserPrincipalName eq '{USER_EMAIL}'").Id
New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId

# Remove user from group
Remove-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId

# Get group members
Get-MgGroupMember -GroupId $groupId

# Delete group
Remove-MgGroup -GroupId $groupId
```

### Application Management
```powershell
# List all applications
Get-MgApplication -All | Select-Object DisplayName, AppId, Id

# Get specific application
Get-MgApplication -Filter "DisplayName eq '{APP_NAME}'"

# Create new application registration
New-MgApplication -DisplayName "{APP_NAME}" `
                  -SignInAudience "AzureADMyOrg"

# Update application
Update-MgApplication -ApplicationId $appId -Web @{ RedirectUris = @("https://app.{TENANT_NAME}.com/callback") }
```

---

## 📂 SharePoint Online Management

### PnP PowerShell (Recommended)
```powershell
# Connect to SharePoint admin center
Connect-PnPOnline -Url "https://{TENANT_NAME}-admin.sharepoint.com" -Interactive

# Connect to specific site
Connect-PnPOnline -Url "{SITE_URL}" -Interactive

# Alternative: Connect with app-only authentication
Connect-PnPOnline -Url "{SITE_URL}" -ClientId "{CLIENT_ID}" -ClientSecret "{CLIENT_SECRET}"

# Get all sites
Get-PnPTenantSite | Select-Object Title, Url, Template

# Create new site collection
New-PnPSite -Type TeamSite -Title "{SITE_TITLE}" -Alias "{SITE_ALIAS}" -Owner "{USER_EMAIL}"

# Get site information
Get-PnPSite

# List all lists/libraries
Get-PnPList

# Create new list
New-PnPList -Title "{LIST_NAME}" -Template GenericList

# Upload file to document library
Add-PnPFile -Path "C:\temp\document.pdf" -Folder "Documents"

# Set site sharing settings
Set-PnPSite -Sharing ExternalUserSharingOnly

# Get site storage usage
Get-PnPSite | Select-Object StorageUsageCurrent, StorageQuota

# Disconnect
Disconnect-PnPOnline
```

### Official SharePoint Module
```powershell
# Connect to SharePoint admin center
Connect-SPOService -Url "https://{TENANT_NAME}-admin.sharepoint.com"

# Get all site collections
Get-SPOSite | Select-Object Title, Url, SharingCapability

# Get specific site
Get-SPOSite -Identity "{SITE_URL}"

# Set site storage quota
Set-SPOSite -Identity "{SITE_URL}" -StorageQuota 2048

# Set sharing capability
Set-SPOSite -Identity "{SITE_URL}" -SharingCapability ExternalUserSharingOnly

# Get tenant settings
Get-SPOTenant

# Disconnect
Disconnect-SPOService
```

---

## 👥 Microsoft Teams Management

### Basic Operations
```powershell
# Connect to Teams
Connect-MicrosoftTeams

# Get all teams
Get-Team | Select-Object DisplayName, GroupId, Visibility

# Get specific team
Get-Team -DisplayName "{TEAM_NAME}"

# Create new team
New-Team -DisplayName "{TEAM_NAME}" -Description "{TEAM_DESCRIPTION}" -Visibility Private

# Add member to team
$teamId = (Get-Team -DisplayName "{TEAM_NAME}").GroupId
Add-TeamUser -GroupId $teamId -User "{USER_EMAIL}" -Role Member

# Add owner to team
Add-TeamUser -GroupId $teamId -User "{USER_EMAIL}" -Role Owner

# Remove user from team
Remove-TeamUser -GroupId $teamId -User "{USER_EMAIL}"

# Get team members
Get-TeamUser -GroupId $teamId

# Create new channel
New-TeamChannel -GroupId $teamId -DisplayName "{CHANNEL_NAME}" -Description "{CHANNEL_DESCRIPTION}"

# Get team channels
Get-TeamChannel -GroupId $teamId

# Archive team
Set-Team -GroupId $teamId -Archived:$true

# Disconnect
Disconnect-MicrosoftTeams
```

### Teams Policies
```powershell
# Get messaging policies
Get-CsTeamsMessagingPolicy

# Create custom messaging policy
New-CsTeamsMessagingPolicy -Identity "RestrictedMessaging" -AllowUrlPreviews $false -AllowUserChat $true

# Assign policy to user
Grant-CsTeamsMessagingPolicy -Identity "{USER_EMAIL}" -PolicyName "RestrictedMessaging"

# Get meeting policies
Get-CsTeamsMeetingPolicy

# Get calling policies
Get-CsTeamsCallingPolicy
```

---

## 📧 Exchange Online & Defender for Office 365

### Basic Connection and Operations
```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Check connection
Get-ConnectionInformation

# Get organization configuration
Get-OrganizationConfig | Select-Object Name, ExchangeVersion

# List all mailboxes
Get-Mailbox -ResultSize 100 | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails

# Get specific mailbox
Get-Mailbox "{USER_EMAIL}"

# Create new mailbox
New-Mailbox -Name "{DISPLAY_NAME}" -UserPrincipalName "{USER_EMAIL}" -Password (ConvertTo-SecureString "TempPass123!" -AsPlainText -Force)

# Create shared mailbox
New-Mailbox -Name "{SHARED_MAILBOX_NAME}" -PrimarySmtpAddress "{MAILBOX_EMAIL}" -Shared

# Add permission to shared mailbox
Add-MailboxPermission -Identity "{SHARED_MAILBOX_EMAIL}" -User "{USER_EMAIL}" -AccessRights FullAccess

# Get mailbox statistics
Get-MailboxStatistics "{USER_EMAIL}" | Select-Object DisplayName, TotalItemSize, ItemCount

# Set mailbox quota
Set-Mailbox "{USER_EMAIL}" -ProhibitSendQuota "2GB" -ProhibitSendReceiveQuota "2.5GB"

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false
```

### Distribution Groups
```powershell
# Get all distribution groups
Get-DistributionGroup | Select-Object DisplayName, PrimarySmtpAddress

# Create new distribution group
New-DistributionGroup -Name "{GROUP_NAME}" -PrimarySmtpAddress "{GROUP_EMAIL}"

# Add member to distribution group
Add-DistributionGroupMember -Identity "{GROUP_EMAIL}" -Member "{USER_EMAIL}"

# Remove member from distribution group
Remove-DistributionGroupMember -Identity "{GROUP_EMAIL}" -Member "{USER_EMAIL}"

# Get group members
Get-DistributionGroupMember -Identity "{GROUP_EMAIL}"
```

### Defender for Office 365
```powershell
# Get Safe Links policies
Get-SafeLinksPolicy

# Create Safe Links policy
New-SafeLinksPolicy -Name "CustomSafeLinks" -EnableSafeLinksForTeams $true -ScanUrls $true

# Get Safe Attachments policies
Get-SafeAttachmentPolicy

# Create Safe Attachments policy
New-SafeAttachmentPolicy -Name "CustomSafeAttachments" -Enable $true -Action Block

# Get Anti-Phishing policies
Get-AntiPhishPolicy

# Get ATP policy for Office 365
Get-ATPPolicyForO365
```

### Mail Flow Rules
```powershell
# Get transport rules
Get-TransportRule | Select-Object Name, State, Priority

# Create new transport rule
New-TransportRule -Name "Block External Auto-Forward" `
                  -FromScope NotInOrganization `
                  -MessageTypeMatches AutoForward `
                  -RejectMessageReasonText "External auto-forwarding is not allowed"

# Enable/disable transport rule
Enable-TransportRule -Identity "Block External Auto-Forward"
Disable-TransportRule -Identity "Block External Auto-Forward"
```

---

## ⚡ Power Platform Management

### Power Apps Administration
```powershell
# Connect to Power Platform
Add-PowerAppsAccount

# Get all environments
Get-PowerAppEnvironment | Select-Object DisplayName, EnvironmentName, Location

# Create new environment
New-PowerAppEnvironment -DisplayName "{ENVIRONMENT_NAME}" -LocationName "unitedstates" -EnvironmentType Sandbox

# Get all Power Apps
Get-PowerApp | Select-Object DisplayName, AppName, Owner

# Get apps in specific environment
Get-PowerApp -EnvironmentName "{ENVIRONMENT_NAME}"

# Remove Power App
Remove-PowerApp -AppName "{APP_ID}"

# Get Power App connections
Get-PowerAppConnection | Select-Object DisplayName, ConnectionName

# Get Power Automate flows
Get-PowerAutomate | Select-Object DisplayName, FlowName, State
```

### Power Platform for Makers
```powershell
# Connect as maker
Add-PowerAppsAccount

# Get my apps
Get-PowerApp | Where-Object {$_.Owner.email -eq "{YOUR_EMAIL}"}

# Get app details
Get-PowerApp -AppName "{APP_ID}"

# Publish app
Publish-PowerApp -AppName "{APP_ID}"

# Set app as featured
Set-PowerAppAsFeatured -AppName "{APP_ID}"
```

---

## 🔧 Microsoft Intune Device Management

### Device Management
```powershell
# Connect to Intune
Connect-MSGraph

# Get all managed devices
Get-IntuneManagedDevice | Select-Object DeviceName, OperatingSystem, ComplianceState, LastSyncDateTime

# Get specific device
Get-IntuneManagedDevice -Filter "DeviceName eq '{DEVICE_NAME}'"

# Get device by user
Get-IntuneManagedDevice -Filter "UserPrincipalName eq '{USER_EMAIL}'"

# Sync device
Invoke-IntuneManagedDeviceSyncDevice -ManagedDeviceId "{DEVICE_ID}"

# Restart device
Invoke-IntuneManagedDeviceRebootNow -ManagedDeviceId "{DEVICE_ID}"

# Wipe device
Invoke-IntuneManagedDeviceWipe -ManagedDeviceId "{DEVICE_ID}"

# Get device compliance
Get-IntuneDeviceCompliancePolicy | Select-Object DisplayName, Platform
```

### Application Management
```powershell
# Get all applications
Get-IntuneApplication | Select-Object DisplayName, Publisher, Platform

# Get application assignments
Get-IntuneApplicationAssignment -ApplicationId "{APP_ID}"

# Get mobile app categories
Get-IntuneMobileAppCategory
```

---

## 🔍 Microsoft Purview Compliance

### eDiscovery and Compliance Search
```powershell
# Connect with compliance scopes
Connect-MgGraph -Scopes "CompliancePolicy.Read.All", "eDiscovery.Read.All"

# Note: Many compliance operations require Exchange Online connection
Connect-ExchangeOnline

# Get compliance searches
Get-ComplianceSearch | Select-Object Name, Status, Items, Size

# Create new compliance search
New-ComplianceSearch -Name "Legal Hold Search - {CASE_NAME}" `
                     -ContentMatchQuery "Subject:'{SEARCH_TERM}'" `
                     -ExchangeLocation All

# Start compliance search
Start-ComplianceSearch -Identity "Legal Hold Search - {CASE_NAME}"

# Get search results
Get-ComplianceSearch -Identity "Legal Hold Search - {CASE_NAME}" | Select-Object Items, Size, Status

# Create compliance search action (export)
New-ComplianceSearchAction -SearchName "Legal Hold Search - {CASE_NAME}" -Export -ShareRootPath "\\server\share"
```

### Retention Policies
```powershell
# Get retention policies
Get-RetentionCompliancePolicy | Select-Object Name, Mode, Enabled

# Create retention policy
New-RetentionCompliancePolicy -Name "{POLICY_NAME}" -ExchangeLocation All -SharePointLocation All

# Get retention rules
Get-RetentionComplianceRule | Select-Object Name, Policy, RetentionDuration
```

---

## 📊 Reporting and Monitoring

### Microsoft Graph Reports
```powershell
# Connect with reports permissions
Connect-MgGraph -Scopes "Reports.Read.All"

# Get Office 365 usage reports
Get-MgReportOffice365ActiveUser -Period D30 | Out-GridView
Get-MgReportOffice365GroupActivity -Period D30 | Out-GridView
Get-MgReportSharePointSiteUsage -Period D30 | Out-GridView

# Get Teams usage
Get-MgReportTeamsUserActivity -Period D30 | Out-GridView
Get-MgReportTeamsDeviceUsage -Period D30 | Out-GridView

# Get Exchange usage
Get-MgReportEmailActivity -Period D30 | Out-GridView
Get-MgReportMailboxUsage -Period D30 | Out-GridView
```

### Security Reports
```powershell
# Connect with security permissions
Connect-MgGraph -Scopes "SecurityEvents.Read.All"

# Get security alerts
Get-MgSecurityAlert -Top 50 | Select-Object Title, Severity, Status, CreatedDateTime

# Get security incidents
Get-MgSecurityIncident | Select-Object DisplayName, Status, Severity, CreatedDateTime
```

---

## 🔄 Bulk Operations

### Bulk User Operations
```powershell
# Import users from CSV
$users = Import-Csv "C:\temp\NewUsers.csv"
foreach ($user in $users) {
    $passwordProfile = @{
        Password = "TempPass123!"
        ForceChangePasswordNextSignIn = $true
    }
    
    New-MgUser -DisplayName $user.DisplayName `
               -UserPrincipalName $user.UserPrincipalName `
               -MailNickname $user.MailNickname `
               -PasswordProfile $passwordProfile `
               -AccountEnabled:$true
}

# Bulk update user properties
$users = Get-MgUser -Filter "Department eq 'Sales'"
foreach ($user in $users) {
    Update-MgUser -UserId $user.Id -CompanyName "Contoso Sales Division"
}
```

### Bulk License Assignment
```powershell
# Get available licenses
Get-MgSubscribedSku | Select-Object SkuPartNumber, ConsumedUnits, PrepaidUnits

# Assign license to multiple users
$licenseSkuId = (Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPREMIUM"}).SkuId
$usersToLicense = Get-MgUser -Filter "Department eq 'IT'"

foreach ($user in $usersToLicense) {
    Set-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId = $licenseSkuId} -RemoveLicenses @()
}
```

---

## 🛠️ Utility Scripts

### Environment Setup
```powershell
# Set variables for your environment
$TenantName = "{TENANT_NAME}"
$TenantDomain = "{TENANT_DOMAIN}"
$AdminEmail = "{ADMIN_EMAIL}"

# Connect to all services
Connect-MgGraph -Scopes "Directory.ReadWrite.All", "User.ReadWrite.All", "Group.ReadWrite.All"
Connect-ExchangeOnline
Connect-MicrosoftTeams
Connect-PnPOnline -Url "https://$TenantName-admin.sharepoint.com" -Interactive

Write-Host "Connected to all Microsoft 365 services for tenant: $TenantDomain" -ForegroundColor Green
```

### Disconnect All Services
```powershell
# Disconnect from all services
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MicrosoftTeams
Disconnect-PnPOnline

Write-Host "Disconnected from all Microsoft 365 services" -ForegroundColor Yellow
```

### Health Check Script
```powershell
# Check service connectivity
function Test-M365Connectivity {
    $results = @()
    
    # Test Microsoft Graph
    try {
        $context = Get-MgContext
        if ($context) {
            $results += "✅ Microsoft Graph: Connected ($($context.Account))"
        } else {
            $results += "❌ Microsoft Graph: Not connected"
        }
    } catch {
        $results += "❌ Microsoft Graph: Error - $($_.Exception.Message)"
    }
    
    # Test Exchange Online
    try {
        $conn = Get-ConnectionInformation
        if ($conn) {
            $results += "✅ Exchange Online: Connected"
        } else {
            $results += "❌ Exchange Online: Not connected"
        }
    } catch {
        $results += "❌ Exchange Online: Not connected"
    }
    
    # Test Teams
    try {
        $tenant = Get-CsTenant -ErrorAction SilentlyContinue
        if ($tenant) {
            $results += "✅ Microsoft Teams: Connected"
        } else {
            $results += "❌ Microsoft Teams: Not connected"
        }
    } catch {
        $results += "❌ Microsoft Teams: Not connected"
    }
    
    return $results
}

# Run connectivity test
Test-M365Connectivity
```

---

## 📝 CSV Templates

### Users Import Template (NewUsers.csv)
```csv
DisplayName,UserPrincipalName,MailNickname,JobTitle,Department,Manager
John Doe,john.doe@{TENANT_DOMAIN},johndoe,Developer,IT,manager@{TENANT_DOMAIN}
Jane Smith,jane.smith@{TENANT_DOMAIN},janesmith,Designer,Marketing,manager@{TENANT_DOMAIN}
```

### Bulk License Assignment Template (UsersLicense.csv)
```csv
UserPrincipalName,LicenseSku,Department
user1@{TENANT_DOMAIN},ENTERPRISEPREMIUM,IT
user2@{TENANT_DOMAIN},ENTERPRISEPREMIUM,Sales
```

---

## 🔧 Advanced Configuration

### Custom PowerShell Profile
Add this to your PowerShell profile (`$PROFILE`) for quick access:

```powershell
# Microsoft 365 Quick Connect Functions
function Connect-M365All {
    param(
        [string]$TenantName = "{TENANT_NAME}"
    )
    
    Write-Host "Connecting to Microsoft 365 services..." -ForegroundColor Cyan
    
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Directory.ReadWrite.All", "User.ReadWrite.All", "Group.ReadWrite.All"
    
    # Connect to Exchange Online
    Connect-ExchangeOnline
    
    # Connect to Teams
    Connect-MicrosoftTeams
    
    # Connect to SharePoint
    Connect-PnPOnline -Url "https://$TenantName-admin.sharepoint.com" -Interactive
    
    Write-Host "✅ Connected to all Microsoft 365 services!" -ForegroundColor Green
}

function Disconnect-M365All {
    Write-Host "Disconnecting from Microsoft 365 services..." -ForegroundColor Yellow
    
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    
    Write-Host "✅ Disconnected from all services!" -ForegroundColor Green
}

# Set aliases for quick access
Set-Alias -Name "m365connect" -Value "Connect-M365All"
Set-Alias -Name "m365disconnect" -Value "Disconnect-M365All"
```

---

## 🚨 Important Notes

1. **Replace all variables** in `{BRACKETS}` with your actual values
2. **Test in development** environment first
3. **Use appropriate permissions** - follow principle of least privilege
4. **Monitor rate limits** when performing bulk operations
5. **Always disconnect** sessions when finished
6. **Keep credentials secure** - never hardcode passwords in scripts
7. **Regular updates** - keep modules updated with `Update-Module`

For more detailed information, refer to the complete usage guide in `README-Microsoft365-PowerShell.md`.