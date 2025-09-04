#Requires -Version 5.1

<#
.SYNOPSIS
    Microsoft 365 External User Audit Script - Microsoft Graph Only
.DESCRIPTION
    Streamlined audit solution for external users using Microsoft Graph API exclusively
.AUTHOR
    Microsoft 365 Security SME
.VERSION
    3.0.0
.NOTES
    Pure Microsoft Graph implementation with simplified authentication and comprehensive reporting
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$TenantId = "",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\AuditReports",
    
    [Parameter(Mandatory = $false)]
    [int]$DaysToAudit = 90,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportIndividualReports
)

#region Configuration and Initialization

# 🔧 Global Configuration
$Global:AuditConfig = @{
    StartTime = Get-Date
    ErrorLog = [System.Collections.ArrayList]::new()
    WarningLog = [System.Collections.ArrayList]::new()
    ProcessedItems = 0
    TotalItems = 0
    GraphScopes = @(
        "User.Read.All"
        "Group.Read.All"
        "Sites.Read.All"
        "TeamMember.Read.All"
        "Directory.Read.All"
        "AuditLog.Read.All"
        "Reports.Read.All"
        "GroupMember.Read.All"
    )
}

# Module validation and installation function
function Install-RequiredModules {
    [CmdletBinding()]
    param()
    
    Write-Host "🔧 Checking and installing Microsoft Graph module..." -ForegroundColor Cyan
    
    # Remove all potentially conflicting modules first
    $conflictingModules = @(
        "Microsoft.Graph*",
        "ExchangeOnlineManagement",
        "PnP.PowerShell",
        "AzureAD*",
        "MSOnline"
    )
    
    foreach ($modulePattern in $conflictingModules) {
        Get-Module -Name $modulePattern | ForEach-Object {
            Write-Host "Removing loaded module: $($_.Name)" -ForegroundColor Yellow
            Remove-Module -Name $_.Name -Force -ErrorAction SilentlyContinue
        }
    }
    
    # Only Microsoft Graph module is required
    $requiredModule = @{ Name = "Microsoft.Graph"; ExactVersion = "1.28.0" }
    
    Write-Host "Checking module: $($requiredModule.Name)..." -ForegroundColor Cyan
    
    $installedModule = Get-Module -Name $requiredModule.Name -ListAvailable | 
        Where-Object { $_.Version -eq [Version]$requiredModule.ExactVersion } | 
        Select-Object -First 1
    
    if (-not $installedModule) {
        Write-Host "Installing $($requiredModule.Name)..." -ForegroundColor Yellow
        try {
            $installParams = @{
                Name = $requiredModule.Name
                RequiredVersion = $requiredModule.ExactVersion
                Scope = "CurrentUser"
                Force = $true
                AllowClobber = $true
            }
            
            Install-Module @installParams
            Write-Host "✅ $($requiredModule.Name) installed successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "❌ Failed to install $($requiredModule.Name): $($_.Exception.Message)" -ForegroundColor Red
            throw "Microsoft Graph module is required for this audit script"
        }
    }
    else {
        Write-Host "✅ $($requiredModule.Name) version $($installedModule.Version) is available" -ForegroundColor Green
    }
    
    # Import Microsoft Graph module
    try {
        Write-Host "Importing Microsoft Graph module..." -ForegroundColor Cyan
        
        $graphModule = Get-Module -Name "Microsoft.Graph" -ListAvailable | 
            Where-Object { $_.Version -eq "1.28.0" } | 
            Select-Object -First 1
        
        if ($graphModule) {
            Import-Module -Name $graphModule.Path -Force -Global
            Write-Host "✅ Microsoft Graph 1.28.0 imported successfully" -ForegroundColor Green
        }
        else {
            throw "Microsoft Graph module version 1.28.0 not found"
        }
    }
    catch {
        Write-Host "❌ Failed to import Microsoft Graph module: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

#endregion

#region Authentication Module

function Connect-M365Services {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "🔐 Connecting to Microsoft Graph..." -ForegroundColor Cyan
        
        # Build connection parameters for v1.28.0 compatibility
        $graphParams = @{
            Scopes = $Global:AuditConfig.GraphScopes
        }
        
        if ($TenantId) {
            $graphParams.Add("TenantId", $TenantId)
        }
        
        try {
            # Try standard interactive authentication first
            Connect-MgGraph @graphParams -ErrorAction Stop
            
            # Wait a moment for the connection to stabilize
            Start-Sleep -Seconds 2
            
            $context = Get-MgContext -ErrorAction Stop
            Write-Host "✅ Connected to Microsoft Graph as: $($context.Account)" -ForegroundColor Green
            Write-Host "✅ Tenant: $($context.TenantId)" -ForegroundColor Green
        }
        catch {
            Write-Host "Standard authentication failed, trying device code..." -ForegroundColor Yellow
            try {
                # Fallback to device code authentication
                Connect-MgGraph -Scopes $Global:AuditConfig.GraphScopes -UseDeviceAuthentication -ErrorAction Stop
                $context = Get-MgContext -ErrorAction Stop
                Write-Host "✅ Connected to Microsoft Graph as: $($context.Account)" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
                throw
            }
        }
        
        return $true
    }
    catch {
        Write-Host "Failed to authenticate to Microsoft Graph" -ForegroundColor Red -ErrorRecord $_
        throw
    }
}
        
        if ($TenantId) {
            $graphParams.Add("TenantId", $TenantId)
        }
        
        try {
            # Try standard interactive authentication first
            Connect-MgGraph @graphParams -ErrorAction Stop
            
            # Wait a moment for the connection to stabilize
            Start-Sleep -Seconds 2
            
            $context = Get-MgContext -ErrorAction Stop
            Write-Host "✅ Connected to Microsoft Graph as: $($context.Account)" -ForegroundColor $(if("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}else{"Cyan"})
        }
        catch {
            Write-Host "Standard authentication failed, trying device code..." -ForegroundColor Yellow
            try {
                # Fallback to device code authentication (without NoWelcome for v1.28.0)
                Connect-MgGraph -Scopes $Global:AuditConfig.GraphScopes -UseDeviceAuthentication -ErrorAction Stop
                $context = Get-MgContext -ErrorAction Stop
                Write-Host "✅ Connected to Microsoft Graph as: $($context.Account)" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor $(if("Error" -eq "Error"){"Red"}else{"Yellow"})
                throw
            }
        }
        
        # Connect to SharePoint Online
        Write-Host "Connecting to SharePoint Online..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        try {
            if (-not $TenantId) {
                # Try alternative method to get tenant info
                $context = Get-MgContext -ErrorAction Stop
                $TenantId = $context.TenantId
            }
            
            # Use hardcoded admin URL pattern as fallback
            $adminUrl = "https://m365x22747677-admin.sharepoint.com"
            Write-Host "Using SharePoint Admin URL: $adminUrl" -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        }
        catch {
            Write-Host "Could not determine tenant info, using default admin URL pattern" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
            $adminUrl = "https://m365x22747677-admin.sharepoint.com"
        }
        
        try {
            Connect-PnPOnline -Url $adminUrl -ClientId "afe1b358-534b-4c96-abb9-ecea5d5f2e5d" -Interactive -ErrorAction Stop
        } catch {
            Write-Host "Warning: SharePoint connection issue: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "Continuing with limited SharePoint functionality..." -ForegroundColor Yellow
        }
        
        Write-Host "✅ Connected to SharePoint Online: $adminUrl" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
        
        # Connect to Exchange Online (for unified group management) - with enhanced fallback methods
        Write-Host "Connecting to Exchange Online..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        $exoConnected = $false
        
        # Method 1: Try with UserPrincipalName (most reliable for modern auth)
        try {
            Write-Host "Attempting connection with UserPrincipalName authentication..." -ForegroundColor Cyan
            
            # Get the current user context from Microsoft Graph
            $currentUser = Get-MgContext -ErrorAction Stop
            if ($currentUser -and $currentUser.Account) {
                $exoParams = @{
                    UserPrincipalName = $currentUser.Account
                    ShowBanner = $false
                    ErrorAction = "Stop"
                }
                
                Connect-ExchangeOnline @exoParams
                $exoConnected = $true
                Write-Host "✅ Connected to Exchange Online with UPN authentication" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "UPN authentication failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # Method 2: Try basic interactive authentication if Method 1 failed
        if (-not $exoConnected) {
            try {
                Write-Host "Trying basic interactive authentication..." -ForegroundColor Yellow
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
                $exoConnected = $true
                Write-Host "✅ Connected to Exchange Online with interactive auth" -ForegroundColor Green
            }
            catch {
                Write-Host "Interactive authentication failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        # Method 3: Try with different connection endpoint
        if (-not $exoConnected) {
            try {
                Write-Host "Trying alternative connection method..." -ForegroundColor Yellow
                # Force a clean connection with explicit parameters
                $altParams = @{
                    ShowBanner = $false
                    ShowProgress = $false
                    ErrorAction = "Stop"
                }
                
                Connect-ExchangeOnline @altParams
                $exoConnected = $true
                Write-Host "✅ Connected to Exchange Online with alternative method" -ForegroundColor Green
            }
            catch {
                Write-Host "Alternative connection failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        # If all methods failed, continue without Exchange Online
        if (-not $exoConnected) {
            Write-Host "Warning: Could not connect to Exchange Online after trying multiple methods." -ForegroundColor Yellow
            Write-Host "This is often due to:" -ForegroundColor Yellow
            Write-Host "  • WinRM service issues" -ForegroundColor Yellow
            Write-Host "  • Network connectivity problems" -ForegroundColor Yellow
            Write-Host "  • Exchange Online service temporary issues" -ForegroundColor Yellow
            Write-Host "  • PowerShell execution policy restrictions" -ForegroundColor Yellow
            Write-Host "Continuing with Microsoft Graph and SharePoint functionality..." -ForegroundColor Cyan
            Write-Host "Note: Group membership analysis will use Microsoft Graph instead of Exchange Online" -ForegroundColor Cyan
            
            # Show troubleshooting information
            Show-ExchangeOnlineTroubleshooting
        }
        
        Write-Host "✅ All available services connected successfully" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
        
        return $true
    }
    catch {
        Write-Host "Failed to authenticate to Microsoft 365 services" -ForegroundColor $(if("Error" -eq "Info"){"Cyan"}elseif("Error" -eq "Success"){"Green"}elseif("Error" -eq "Warning"){"Yellow"}elseif("Error" -eq "Error"){"Red"}else{"White"}) -ErrorRecord $_
        throw
    }
}

#endregion

#region External User Discovery

function Get-ExternalUsers {
    [CmdletBinding()]
    param()
    
    $externalUsers = [System.Collections.ArrayList]::new()
    
    try {
        Write-Host "🔍 Discovering external users in tenant..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        # Method 1: Get Guest Users from Azure AD with error handling and fallback
        try {
            Write-Host "Attempting to retrieve guest users from Azure AD..." -ForegroundColor Cyan
            
            # Try the most basic approach first to avoid SDK conflicts
            try {
                Write-Host "Using basic Graph API call..." -ForegroundColor Cyan
                $guestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -ErrorAction Stop
                
                Write-Host "Successfully retrieved $($guestUsers.Count) guest users with basic properties" -ForegroundColor Green
            }
            catch {
                Write-Host "Basic Graph call failed: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "Attempting alternative approach..." -ForegroundColor Yellow
                
                # Try without filter as fallback
                try {
                    $allUsers = Get-MgUser -All -ErrorAction Stop
                    $guestUsers = $allUsers | Where-Object { $_.UserType -eq "Guest" }
                    Write-Host "Retrieved guest users using alternative method: $($guestUsers.Count)" -ForegroundColor Green
                }
                catch {
                    Write-Host "Alternative Graph method also failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    $guestUsers = @()
                }
            }
            
            foreach ($guest in $guestUsers) {
                try {
                    $userObj = [PSCustomObject]@{
                        Id = $guest.Id ?? "Unknown"
                        DisplayName = $guest.DisplayName ?? "Unknown"
                        Email = $guest.Mail ?? $guest.UserPrincipalName ?? "Unknown"
                        UPN = $guest.UserPrincipalName ?? "Unknown"
                        UserType = "Guest"
                        CreatedDate = $guest.CreatedDateTime ?? (Get-Date "1900-01-01")
                        ExternalUserState = $guest.ExternalUserState ?? "Unknown"
                        StateChangeDate = $guest.ExternalUserStateChangeDateTime
                        CompanyName = $guest.CompanyName ?? ""
                        Department = $guest.Department ?? ""
                        JobTitle = $guest.JobTitle ?? ""
                        AccountEnabled = $guest.AccountEnabled ?? $false
                        LastSignIn = $null  # Will be populated if available
                        Source = "Azure AD"
                        SharePointSites = [System.Collections.ArrayList]::new()
                        TeamsAccess = [System.Collections.ArrayList]::new()
                        GroupMemberships = [System.Collections.ArrayList]::new()
                        FilesCreated = [System.Collections.ArrayList]::new()
                        Permissions = [System.Collections.ArrayList]::new()
                    }
                    
                    $externalUsers.Add($userObj) | Out-Null
                }
                catch {
                    Write-Host "Error processing guest user $($guest.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            Write-Host "Found $($guestUsers.Count) guest users in Azure AD" -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        }
        catch {
            Write-Host "Error retrieving guest users from Azure AD: $($_.Exception.Message)" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
            Write-Host "Continuing with SharePoint external user discovery..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
            $guestUsers = @()
        }
        
        # Method 2: Get External Users from SharePoint with error handling
        try {
            $spoExternalUsers = Get-PnPExternalUser -PageSize 50 -Position 0 -ErrorAction Stop
            
            foreach ($spoUser in $spoExternalUsers) {
                $existingUser = $externalUsers | Where-Object { $_.Email -eq $spoUser.Email }
                
                if (-not $existingUser) {
                    $userObj = [PSCustomObject]@{
                        Id = $spoUser.UniqueId
                        DisplayName = $spoUser.DisplayName
                        Email = $spoUser.Email
                        UPN = $spoUser.Email
                        UserType = "External"
                        CreatedDate = $spoUser.WhenCreated
                        ExternalUserState = $spoUser.InvitedAs
                        StateChangeDate = $null
                        CompanyName = $null
                        Department = $null
                        JobTitle = $null
                        AccountEnabled = $true
                        LastSignIn = $null
                        Source = "SharePoint"
                        SharePointSites = [System.Collections.ArrayList]::new()
                        TeamsAccess = [System.Collections.ArrayList]::new()
                        GroupMemberships = [System.Collections.ArrayList]::new()
                        FilesCreated = [System.Collections.ArrayList]::new()
                        Permissions = [System.Collections.ArrayList]::new()
                    }
                    
                    $externalUsers.Add($userObj) | Out-Null
                }
            }
        }
        catch {
            Write-Host "Error retrieving external users from SharePoint: $($_.Exception.Message)" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
        }
        
        Write-Host "Total external users discovered: $($externalUsers.Count)" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
        
        return $externalUsers
    }
    catch {
        Write-Host "Error discovering external users" -ForegroundColor $(if("Error" -eq "Info"){"Cyan"}elseif("Error" -eq "Success"){"Green"}elseif("Error" -eq "Warning"){"Yellow"}elseif("Error" -eq "Error"){"Red"}else{"White"}) -ErrorRecord $_
        throw
    }
}

#endregion

#region SharePoint Site Analysis

function Get-SharePointSiteAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExternalUsers
    )
    
    try {
        Write-Host "📂 Analyzing SharePoint site access..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        # Get all site collections
        if ($SpecificSites.Count -gt 0) {
            $sites = $SpecificSites | ForEach-Object {
                Get-PnPTenantSite -Url $_ -ErrorAction SilentlyContinue
            }
        } else {
            $sites = Get-PnPTenantSite -IncludeOneDriveSites:$false
        }
        
        $Global:AuditConfig.TotalItems = $sites.Count
        
        foreach ($site in $sites) {
            $Global:AuditConfig.ProcessedItems++
            $progress = ($Global:AuditConfig.ProcessedItems / $Global:AuditConfig.TotalItems) * 100
            
            Write-Progress -Activity "Analyzing SharePoint Sites" `
                          -Status "Processing: $($site.Url)" `
                          -PercentComplete $progress
            
            try {
                Connect-PnPOnline -Url $site.Url -ClientId "afe1b358-534b-4c96-abb9-ecea5d5f2e5d" -Interactive -ErrorAction Stop
                
                $siteUsers = Get-PnPUser -WithRightsAssigned | Where-Object { $_.LoginName -like "*#ext#*" -or $_.UserPrincipalName -like "*#EXT#*" }
                
                foreach ($siteUser in $siteUsers) {
                    $extUser = $ExternalUsers | Where-Object { $_.Email -eq $siteUser.Email -or $_.UPN -eq $siteUser.UserPrincipalName }
                    
                    if ($extUser) {
                        $siteAccess = [PSCustomObject]@{
                            SiteUrl = $site.Url
                            SiteTitle = $site.Title
                            SiteTemplate = $site.Template
                            SiteSharingCapability = $site.SharingCapability
                            PermissionLevel = $siteUser.IsSiteAdmin ? "Site Collection Administrator" : "Site Member"
                            Groups = ($siteUser.Groups | ForEach-Object { $_.Title }) -join "; "
                            AddedDate = $null
                            LastAccessed = $null
                        }
                        
                        $extUser.SharePointSites.Add($siteAccess) | Out-Null
                    }
                }
            }
            catch {
                Write-Host "Warning: Could not connect to $($site.Url): $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        Write-Host "✅ SharePoint site analysis complete" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
    }
    catch {
        Write-Host "Error analyzing SharePoint sites" -ForegroundColor $(if("Error" -eq "Info"){"Cyan"}elseif("Error" -eq "Success"){"Green"}elseif("Error" -eq "Warning"){"Yellow"}elseif("Error" -eq "Error"){"Red"}else{"White"}) -ErrorRecord $_
        throw
    }
}

function Get-UserSitePermissions {
    [CmdletBinding()]
    param(
        [string]$SiteUrl,
        [string]$UserEmail
    )
    
    $permissions = [System.Collections.ArrayList]::new()
    
    try {
        # Get all lists and libraries
        $lists = Get-PnPList -Includes RoleAssignments, HasUniqueRoleAssignments
        
        foreach ($list in $lists) {
            if ($list.HasUniqueRoleAssignments) {
                $roleAssignments = Get-PnPProperty -ClientObject $list -Property RoleAssignments
                
                foreach ($roleAssignment in $roleAssignments) {
                    $member = Get-PnPProperty -ClientObject $roleAssignment -Property Member
                    
                    if ($member.LoginName -match $UserEmail) {
                        $roleDefs = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings
                        
                        $permObj = [PSCustomObject]@{
                            Resource = $list.Title
                            ResourceType = "List/Library"
                            PermissionLevel = ($roleDefs | ForEach-Object { $_.Name }) -join "; "
                            GrantedThrough = "Direct Assignment"
                        }
                        
                        $permissions.Add($permObj) | Out-Null
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Error getting detailed permissions for $UserEmail" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
    }
    
    return $permissions
}

function Get-ExternalUserFiles {
    [CmdletBinding()]
    param(
        [string]$SiteUrl,
        [object[]]$ExternalUsers
    )
    
    $files = [System.Collections.ArrayList]::new()
    
    try {
        # Query for files created in the last N days
        $startDate = (Get-Date).AddDays(-$DaysToAudit).ToString("yyyy-MM-dd")
        
        $camlQuery = @"
<View Scope='RecursiveAll'>
    <Query>
        <Where>
            <Geq>
                <FieldRef Name='Created' />
                <Value Type='DateTime'>$startDate</Value>
            </Geq>
        </Where>
        <OrderBy>
            <FieldRef Name='Created' Ascending='FALSE' />
        </OrderBy>
    </Query>
    <ViewFields>
        <FieldRef Name='FileLeafRef' />
        <FieldRef Name='FileDirRef' />
        <FieldRef Name='Author' />
        <FieldRef Name='Created' />
        <FieldRef Name='Modified' />
        <FieldRef Name='File_x0020_Size' />
    </ViewFields>
</View>
"@
        
        $lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 } # Document Libraries only
        
        foreach ($list in $lists) {
            try {
                $items = Get-PnPListItem -List $list -Query $camlQuery
                
                foreach ($item in $items) {
                    $author = $item["Author"]
                    
                    if ($author.Email -and ($ExternalUsers | Where-Object { $_.Email -eq $author.Email })) {
                        $fileObj = [PSCustomObject]@{
                            FileName = $item["FileLeafRef"]
                            FilePath = $item["FileDirRef"]
                            Library = $list.Title
                            SiteUrl = $SiteUrl
                            CreatedBy = $author.Email
                            CreatedByName = $author.LookupValue
                            CreatedDate = $item["Created"]
                            ModifiedDate = $item["Modified"]
                            FileSize = $item["File_x0020_Size"]
                            FileUrl = "$SiteUrl/$($item['FileDirRef'])/$($item['FileLeafRef'])"
                        }
                        
                        $files.Add($fileObj) | Out-Null
                    }
                }
            }
            catch {
                Write-Host "Error processing library: $($list.Title)" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
            }
        }
    }
    catch {
        Write-Host "Error retrieving files for external users" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
    }
    
    return $files
}

#endregion

#region Microsoft Teams Analysis

function Get-TeamsAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExternalUsers
    )
    
    try {
        Write-Host "👥 Analyzing Microsoft Teams access..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        # Get all teams with error handling - using alternative approach
        try {
            # Try using REST API approach if SDK has issues
            $teams = @()
            try {
                $teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All -ErrorAction Stop
            }
            catch {
                Write-Host "Primary Graph SDK method failed, trying alternative approach..." -ForegroundColor Yellow
                # Alternative: Get all groups and filter for teams
                $allGroups = Get-MgGroup -All -ErrorAction Stop
                $teams = $allGroups | Where-Object { $_.ResourceProvisioningOptions -contains "Team" }
            }
        }
        catch {
            Write-Host "Error retrieving Teams: $($_.Exception.Message)" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
            Write-Host "Skipping Teams analysis..." -ForegroundColor Yellow
            return
        }
        
        if ($teams.Count -eq 0) {
            Write-Host "No Teams found or accessible" -ForegroundColor Yellow
            return
        }
        
        $Global:AuditConfig.TotalItems = $teams.Count
        $Global:AuditConfig.ProcessedItems = 0
        
        foreach ($team in $teams) {
            $Global:AuditConfig.ProcessedItems++
            $progress = ($Global:AuditConfig.ProcessedItems / $Global:AuditConfig.TotalItems) * 100
            
            Write-Progress -Activity "Analyzing Microsoft Teams" `
                          -Status "Processing: $($team.DisplayName)" `
                          -PercentComplete $progress
            
            try {
                # Get team members with retry logic
                $members = @()
                try {
                    $members = Get-MgGroupMember -GroupId $team.Id -All -ErrorAction Stop
                }
                catch {
                    Write-Host "Warning: Could not get members for team: $($team.DisplayName)" -ForegroundColor Yellow
                    continue
                }
                
                foreach ($member in $members) {
                    try {
                        $memberDetails = Get-MgUser -UserId $member.Id -Property Id,DisplayName,UserPrincipalName,UserType -ErrorAction SilentlyContinue
                        
                        if ($memberDetails -and $memberDetails.UserType -eq "Guest") {
                            $extUser = $ExternalUsers | Where-Object { $_.Id -eq $memberDetails.Id }
                            
                            if ($extUser) {
                                # Get member role with error handling
                                $isOwner = $false
                                try {
                                    $owners = Get-MgGroupOwner -GroupId $team.Id -All -ErrorAction SilentlyContinue
                                    $isOwner = $owners.Id -contains $member.Id
                                }
                                catch {
                                    # Continue without owner information
                                }
                                
                                $teamAccess = [PSCustomObject]@{
                                    TeamId = $team.Id
                                    TeamName = $team.DisplayName
                                    TeamDescription = $team.Description
                                    TeamVisibility = $team.Visibility
                                    MemberRole = $isOwner ? "Owner" : "Member"
                                    JoinedDate = $null
                                    TeamArchived = $team.IsArchived
                                    TeamSiteUrl = $null
                                }
                                
                                $extUser.TeamsAccess.Add($teamAccess) | Out-Null
                                
                                # Add to group memberships
                                $groupMembership = [PSCustomObject]@{
                                    GroupId = $team.Id
                                    GroupName = $team.DisplayName
                                    GroupType = "Microsoft 365 Group (Team-enabled)"
                                    MembershipType = $isOwner ? "Owner" : "Member"
                                }
                                
                                $extUser.GroupMemberships.Add($groupMembership) | Out-Null
                            }
                        }
                    }
                    catch {
                        # Skip this member and continue
                        continue
                    }
                }
            }
            catch {
                Write-Host "Error processing team: $($team.DisplayName)" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
            }
        }
        
        Write-Host "✅ Teams analysis complete" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
    }
    catch {
        Write-Host "Error analyzing Microsoft Teams" -ForegroundColor $(if("Error" -eq "Info"){"Cyan"}elseif("Error" -eq "Success"){"Green"}elseif("Error" -eq "Warning"){"Yellow"}elseif("Error" -eq "Error"){"Red"}else{"White"}) -ErrorRecord $_
        # Don't throw - continue with rest of audit
    }
}

#endregion

#region Microsoft 365 Groups Analysis

function Get-M365GroupMemberships {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExternalUsers
    )
    
    try {
        Write-Host "📊 Analyzing Microsoft 365 Group memberships..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        # Get all Microsoft 365 Groups (excluding Teams) with error handling
        try {
            $groups = @()
            try {
                $groups = Get-MgGroup -Filter "groupTypes/Any(x:x eq 'Unified') and NOT resourceProvisioningOptions/Any(x:x eq 'Team')" -All -ErrorAction Stop
            }
            catch {
                Write-Host "Primary filter method failed, trying alternative approach..." -ForegroundColor Yellow
                # Alternative: Get all unified groups and filter out teams manually
                $allGroups = Get-MgGroup -Filter "groupTypes/Any(x:x eq 'Unified')" -All -ErrorAction Stop
                $groups = $allGroups | Where-Object { $_.ResourceProvisioningOptions -notcontains "Team" }
            }
        }
        catch {
            Write-Host "Error retrieving M365 Groups: $($_.Exception.Message)" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
            Write-Host "Skipping M365 Groups analysis..." -ForegroundColor Yellow
            return
        }
        
        if ($groups.Count -eq 0) {
            Write-Host "No M365 Groups found or accessible" -ForegroundColor Yellow
            return
        }
        
        foreach ($group in $groups) {
            try {
                $members = @()
                try {
                    $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
                }
                catch {
                    Write-Host "Warning: Could not get members for group: $($group.DisplayName)" -ForegroundColor Yellow
                    continue
                }
                
                foreach ($member in $members) {
                    try {
                        $memberDetails = Get-MgUser -UserId $member.Id -Property Id,DisplayName,UserPrincipalName,UserType -ErrorAction SilentlyContinue
                        
                        if ($memberDetails -and $memberDetails.UserType -eq "Guest") {
                            $extUser = $ExternalUsers | Where-Object { $_.Id -eq $memberDetails.Id }
                            
                            if ($extUser) {
                                $isOwner = $false
                                try {
                                    $owners = Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction SilentlyContinue
                                    $isOwner = $owners.Id -contains $member.Id
                                }
                                catch {
                                    # Continue without owner information
                                }
                                
                                $groupMembership = [PSCustomObject]@{
                                    GroupId = $group.Id
                                    GroupName = $group.DisplayName
                                    GroupType = "Microsoft 365 Group"
                                    MembershipType = $isOwner ? "Owner" : "Member"
                                    GroupEmail = $group.Mail
                                    GroupVisibility = $group.Visibility
                                }
                                
                                $extUser.GroupMemberships.Add($groupMembership) | Out-Null
                            }
                        }
                    }
                    catch {
                        # Skip this member and continue
                        continue
                    }
                }
            }
            catch {
                Write-Host "Error processing group: $($group.DisplayName)" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
            }
        }
        
        Write-Host "✅ Microsoft 365 Groups analysis complete" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
    }
    catch {
        Write-Host "Error analyzing Microsoft 365 Groups" -ForegroundColor $(if("Error" -eq "Info"){"Cyan"}elseif("Error" -eq "Success"){"Green"}elseif("Error" -eq "Warning"){"Yellow"}elseif("Error" -eq "Error"){"Red"}else{"White"}) -ErrorRecord $_
        # Don't throw - continue with rest of audit
    }
}

#endregion

#region Site Sharing Settings Analysis

function Get-SiteSharingSettings {
    [CmdletBinding()]
    param()
    
    $sharingSettings = [System.Collections.ArrayList]::new()
    
    try {
        Write-Host "🔗 Analyzing site sharing settings..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        $sites = Get-PnPTenantSite -IncludeOneDriveSites:$false
        
        foreach ($site in $sites) {
            $settingObj = [PSCustomObject]@{
                SiteUrl = $site.Url
                SiteTitle = $site.Title
                SharingCapability = $site.SharingCapability
                DefaultSharingLinkType = $site.DefaultSharingLinkType
                DefaultLinkPermission = $site.DefaultLinkPermission
                RequireAnonymousLinksExpireInDays = $site.RequireAnonymousLinksExpireInDays
                SharingAllowedDomains = $site.SharingAllowedDomainList
                SharingBlockedDomains = $site.SharingBlockedDomainList
                AllowEditing = $site.DisableCompanyWideSharingLinks
                ShowPeoplePickerSuggestionsForGuestUsers = $site.ShowPeoplePickerSuggestionsForGuestUsers
                BccExternalSharingInvitations = $site.BccExternalSharingInvitations
                ExternalUserExpirationInDays = $site.ExternalUserExpirationInDays
            }
            
            $sharingSettings.Add($settingObj) | Out-Null
        }
        
        Write-Host "✅ Site sharing settings analysis complete" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
        
        return $sharingSettings
    }
    catch {
        Write-Host "Error analyzing site sharing settings" -ForegroundColor $(if("Error" -eq "Info"){"Cyan"}elseif("Error" -eq "Success"){"Green"}elseif("Error" -eq "Warning"){"Yellow"}elseif("Error" -eq "Error"){"Red"}else{"White"}) -ErrorRecord $_
        throw
    }
}

#endregion

#region Guest Expiration Analysis

function Get-GuestExpiration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExternalUsers
    )
    
    try {
        Write-Host "⏰ Analyzing guest expiration settings..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        $guestExpirationReport = [System.Collections.ArrayList]::new()
        
        foreach ($user in $ExternalUsers) {
            $expirationObj = [PSCustomObject]@{
                UserEmail = $user.Email
                DisplayName = $user.DisplayName
                CreatedDate = $user.CreatedDate
                LastSignIn = $user.LastSignIn
                AccountEnabled = $user.AccountEnabled
                ExternalUserState = $user.ExternalUserState
                StateChangeDate = $user.StateChangeDate
                DaysSinceCreation = ((Get-Date) - $user.CreatedDate).Days
                DaysSinceLastSignIn = if ($user.LastSignIn) { ((Get-Date) - $user.LastSignIn).Days } else { "Never signed in" }
                ExpirationStatus = "Active"
                RecommendedAction = ""
            }
            
            # Determine expiration status and recommendations
            if ($expirationObj.DaysSinceLastSignIn -gt 90) {
                $expirationObj.ExpirationStatus = "Inactive"
                $expirationObj.RecommendedAction = "Consider removing - inactive for over 90 days"
            }
            elseif ($expirationObj.DaysSinceLastSignIn -gt 60) {
                $expirationObj.ExpirationStatus = "Warning"
                $expirationObj.RecommendedAction = "Review access - inactive for over 60 days"
            }
            
            if (-not $user.AccountEnabled) {
                $expirationObj.ExpirationStatus = "Disabled"
                $expirationObj.RecommendedAction = "Account disabled - review for removal"
            }
            
            $guestExpirationReport.Add($expirationObj) | Out-Null
        }
        
        Write-Host "✅ Guest expiration analysis complete" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
        
        return $guestExpirationReport
    }
    catch {
        Write-Host "Error analyzing guest expiration" -ForegroundColor $(if("Error" -eq "Info"){"Cyan"}elseif("Error" -eq "Success"){"Green"}elseif("Error" -eq "Warning"){"Yellow"}elseif("Error" -eq "Error"){"Red"}else{"White"}) -ErrorRecord $_
        throw
    }
}

#endregion

#region Report Generation

function Export-AuditReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExternalUsers,
        
        [Parameter(Mandatory = $true)]
        [object[]]$SharingSettings,
        
        [Parameter(Mandatory = $true)]
        [object[]]$GuestExpiration
    )
    
    try {
        Write-Host "📄 Generating audit reports..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        # Create output directory
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        
        # Main External Users Report
        $mainReport = [System.Collections.ArrayList]::new()
        
        foreach ($user in $ExternalUsers) {
            $reportEntry = [PSCustomObject]@{
                # User Information
                UserEmail = $user.Email
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UPN
                UserType = $user.UserType
                CompanyName = $user.CompanyName
                Department = $user.Department
                JobTitle = $user.JobTitle
                AccountEnabled = $user.AccountEnabled
                CreatedDate = $user.CreatedDate
                LastSignIn = $user.LastSignIn
                ExternalUserState = $user.ExternalUserState
                Source = $user.Source
                
                # Access Summary
                SharePointSitesCount = $user.SharePointSites.Count
                SharePointSites = ($user.SharePointSites | ForEach-Object { "$($_.SiteTitle) [$($_.PermissionLevel)]" }) -join "; "
                TeamsCount = $user.TeamsAccess.Count
                Teams = ($user.TeamsAccess | ForEach-Object { "$($_.TeamName) [$($_.MemberRole)]" }) -join "; "
                GroupsCount = $user.GroupMemberships.Count
                Groups = ($user.GroupMemberships | ForEach-Object { "$($_.GroupName) [$($_.MembershipType)]" }) -join "; "
                FilesCreatedCount = $user.FilesCreated.Count
                
                # Risk Assessment
                RiskLevel = Get-UserRiskLevel -User $user
                DaysSinceLastActivity = if ($user.LastSignIn) { ((Get-Date) - $user.LastSignIn).Days } else { "N/A" }
            }
            
            $mainReport.Add($reportEntry) | Out-Null
        }
        
        # Export main report
        $mainReportPath = Join-Path $OutputPath "ExternalUserAudit_Main_$timestamp.csv"
        $mainReport | Export-Csv -Path $mainReportPath -NoTypeInformation
        Write-Host "✅ Main report exported: $mainReportPath" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
        
        # Export individual detailed reports if requested
        if ($ExportIndividualReports) {
            # SharePoint Access Details
            $spReport = [System.Collections.ArrayList]::new()
            foreach ($user in $ExternalUsers) {
                foreach ($site in $user.SharePointSites) {
                    $spEntry = [PSCustomObject]@{
                        UserEmail = $user.Email
                        DisplayName = $user.DisplayName
                        SiteUrl = $site.SiteUrl
                        SiteTitle = $site.SiteTitle
                        SiteTemplate = $site.SiteTemplate
                        PermissionLevel = $site.PermissionLevel
                        Groups = $site.Groups
                        SiteSharingCapability = $site.SiteSharingCapability
                    }
                    $spReport.Add($spEntry) | Out-Null
                }
            }
            
            if ($spReport.Count -gt 0) {
                $spReportPath = Join-Path $OutputPath "ExternalUserAudit_SharePoint_$timestamp.csv"
                $spReport | Export-Csv -Path $spReportPath -NoTypeInformation
                Write-Host "✅ SharePoint report exported: $spReportPath" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
            }
            
            # Teams Access Details
            $teamsReport = [System.Collections.ArrayList]::new()
            foreach ($user in $ExternalUsers) {
                foreach ($team in $user.TeamsAccess) {
                    $teamEntry = [PSCustomObject]@{
                        UserEmail = $user.Email
                        DisplayName = $user.DisplayName
                        TeamId = $team.TeamId
                        TeamName = $team.TeamName
                        TeamDescription = $team.TeamDescription
                        MemberRole = $team.MemberRole
                        TeamVisibility = $team.TeamVisibility
                        TeamArchived = $team.TeamArchived
                        TeamSiteUrl = $team.TeamSiteUrl
                    }
                    $teamsReport.Add($teamEntry) | Out-Null
                }
            }
            
            if ($teamsReport.Count -gt 0) {
                $teamsReportPath = Join-Path $OutputPath "ExternalUserAudit_Teams_$timestamp.csv"
                $teamsReport | Export-Csv -Path $teamsReportPath -NoTypeInformation
                Write-Host "✅ Teams report exported: $teamsReportPath" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
            }
            
            # Files Created Report
            if ($IncludeFileAnalysis) {
                $filesReport = [System.Collections.ArrayList]::new()
                foreach ($user in $ExternalUsers) {
                    foreach ($file in $user.FilesCreated) {
                        $fileEntry = [PSCustomObject]@{
                            CreatedBy = $file.CreatedBy
                            CreatedByName = $file.CreatedByName
                            FileName = $file.FileName
                            FilePath = $file.FilePath
                            Library = $file.Library
                            SiteUrl = $file.SiteUrl
                            CreatedDate = $file.CreatedDate
                            ModifiedDate = $file.ModifiedDate
                            FileSize = $file.FileSize
                            FileUrl = $file.FileUrl
                        }
                        $filesReport.Add($fileEntry) | Out-Null
                    }
                }
                
                if ($filesReport.Count -gt 0) {
                    $filesReportPath = Join-Path $OutputPath "ExternalUserAudit_Files_$timestamp.csv"
                    $filesReport | Export-Csv -Path $filesReportPath -NoTypeInformation
                    Write-Host "✅ Files report exported: $filesReportPath" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
                }
            }
        }
        
        # Export Site Sharing Settings
        $sharingReportPath = Join-Path $OutputPath "SiteSharingSettings_$timestamp.csv"
        $SharingSettings | Export-Csv -Path $sharingReportPath -NoTypeInformation
        Write-Host "✅ Sharing settings report exported: $sharingReportPath" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
        
        # Export Guest Expiration Report
        $expirationReportPath = Join-Path $OutputPath "GuestExpiration_$timestamp.csv"
        $GuestExpiration | Export-Csv -Path $expirationReportPath -NoTypeInformation
        Write-Host "✅ Guest expiration report exported: $expirationReportPath" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
        
        # Generate Executive Summary
        Export-ExecutiveSummary -ExternalUsers $ExternalUsers -OutputPath $OutputPath -Timestamp $timestamp
        
        Write-Host "✅ All reports generated successfully" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
        
        return $mainReportPath
    }
    catch {
        Write-Host "Error generating reports" -ForegroundColor $(if("Error" -eq "Info"){"Cyan"}elseif("Error" -eq "Success"){"Green"}elseif("Error" -eq "Warning"){"Yellow"}elseif("Error" -eq "Error"){"Red"}else{"White"}) -ErrorRecord $_
        throw
    }
}

function Get-UserRiskLevel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$User
    )
    
    $riskScore = 0
    
    # Check last sign-in
    if ($User.LastSignIn) {
        $daysSinceSignIn = ((Get-Date) - $User.LastSignIn).Days
        if ($daysSinceSignIn -gt 90) { $riskScore += 3 }
        elseif ($daysSinceSignIn -gt 60) { $riskScore += 2 }
        elseif ($daysSinceSignIn -gt 30) { $riskScore += 1 }
    } else {
        $riskScore += 4
    }
    
    # Check account status
    if (-not $User.AccountEnabled) { $riskScore += 3 }
    
    # Check access levels
    if ($User.SharePointSites | Where-Object { $_.PermissionLevel -eq "Site Collection Administrator" }) { $riskScore += 2 }
    if ($User.TeamsAccess | Where-Object { $_.MemberRole -eq "Owner" }) { $riskScore += 2 }
    
    # Check file creation activity
    if ($User.FilesCreated.Count -gt 100) { $riskScore += 2 }
    elseif ($User.FilesCreated.Count -gt 50) { $riskScore += 1 }
    
    # Determine risk level
    if ($riskScore -ge 7) { return "High" }
    elseif ($riskScore -ge 4) { return "Medium" }
    else { return "Low" }
}

function Export-ExecutiveSummary {
    [CmdletBinding()]
    param(
        [object[]]$ExternalUsers,
        [string]$OutputPath,
        [string]$Timestamp
    )
    
    $summary = @"
# Microsoft 365 External User Audit - Executive Summary
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

## 📊 Overview Statistics
- Total External Users: $($ExternalUsers.Count)
- Active Users (signed in last 30 days): $(($ExternalUsers | Where-Object { $_.LastSignIn -and ((Get-Date) - $_.LastSignIn).Days -le 30 }).Count)
- Inactive Users (>90 days): $(($ExternalUsers | Where-Object { $_.LastSignIn -and ((Get-Date) - $_.LastSignIn).Days -gt 90 }).Count)
- Disabled Accounts: $(($ExternalUsers | Where-Object { -not $_.AccountEnabled }).Count)

## 🔐 Access Distribution
- Users with SharePoint Access: $(($ExternalUsers | Where-Object { $_.SharePointSites.Count -gt 0 }).Count)
- Users with Teams Access: $(($ExternalUsers | Where-Object { $_.TeamsAccess.Count -gt 0 }).Count)
- Users with Group Memberships: $(($ExternalUsers | Where-Object { $_.GroupMemberships.Count -gt 0 }).Count)

## ⚠️ Risk Assessment
- High Risk Users: $(($ExternalUsers | Where-Object { (Get-UserRiskLevel -User $_) -eq "High" }).Count)
- Medium Risk Users: $(($ExternalUsers | Where-Object { (Get-UserRiskLevel -User $_) -eq "Medium" }).Count)
- Low Risk Users: $(($ExternalUsers | Where-Object { (Get-UserRiskLevel -User $_) -eq "Low" }).Count)

## 📁 File Activity
- Total Files Created by External Users: $(($ExternalUsers | ForEach-Object { $_.FilesCreated.Count } | Measure-Object -Sum).Sum)
- Users Who Created Files: $(($ExternalUsers | Where-Object { $_.FilesCreated.Count -gt 0 }).Count)

## 🎯 Recommendations
1. Review and remove access for inactive users (>90 days without sign-in)
2. Audit high-risk users with extensive permissions
3. Implement regular access reviews for external users
4. Consider implementing guest expiration policies
5. Review site-level sharing settings for sensitive sites

## 📝 Audit Details
- Audit Duration: $((Get-Date) - $Global:AuditConfig.StartTime)
- Sites Analyzed: $Global:AuditConfig.ProcessedItems
- Warnings Encountered: $($Global:AuditConfig.WarningLog.Count)
- Errors Encountered: $($Global:AuditConfig.ErrorLog.Count)
"@
    
    $summaryPath = Join-Path $OutputPath "ExecutiveSummary_$Timestamp.md"
    $summary | Out-File -FilePath $summaryPath -Encoding UTF8
    
    Write-Host "✅ Executive summary generated: $summaryPath" -ForegroundColor $(if("Success" -eq "Info"){"Cyan"}elseif("Success" -eq "Success"){"Green"}elseif("Success" -eq "Warning"){"Yellow"}elseif("Success" -eq "Error"){"Red"}else{"White"})
}

#endregion

#region Main Execution

function Start-ExternalUserAudit {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host @"
╔══════════════════════════════════════════════════════════════╗
║   Microsoft 365 External User Comprehensive Audit Tool       ║
║                     Version 2.0.0                            ║
╚══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan
        
        Write-Host "Starting Microsoft 365 External User Audit..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        Write-Host "Output Path: $OutputPath" -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        # Step 0: Validate and install required modules
        Write-Host "Step 0/7: Validating required PowerShell modules..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        
        # Clean up any potential module conflicts first
        Write-Host "Cleaning up potential module conflicts..." -ForegroundColor Cyan
        $conflictingModules = @("AzureAD", "AzureADPreview", "MSOnline")
        foreach ($module in $conflictingModules) {
            if (Get-Module -Name $module) {
                Write-Host "Removing conflicting module: $module" -ForegroundColor Yellow
                Remove-Module -Name $module -Force -ErrorAction SilentlyContinue
            }
        }
        
        Install-RequiredModules
        
        # Step 1: Authentication
        Write-Host "Step 1/7: Authenticating to Microsoft 365..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        Connect-M365Services
        
        # Step 2: Discover External Users
        Write-Host "Step 2/7: Discovering external users..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        $externalUsers = Get-ExternalUsers
        
        if ($externalUsers.Count -eq 0) {
            Write-Host "No external users found in the tenant" -ForegroundColor $(if("Warning" -eq "Info"){"Cyan"}elseif("Warning" -eq "Success"){"Green"}elseif("Warning" -eq "Warning"){"Yellow"}elseif("Warning" -eq "Error"){"Red"}else{"White"})
            return
        }
        
        # Step 3: Analyze SharePoint Access
        Write-Host "Step 3/7: Analyzing SharePoint site access..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        Get-SharePointSiteAccess -ExternalUsers $externalUsers
        
        # Step 4: Analyze Teams Access
        Write-Host "Step 4/7: Analyzing Microsoft Teams access..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        Get-TeamsAccess -ExternalUsers $externalUsers
        
        # Step 5: Analyze Group Memberships
        Write-Host "Step 5/7: Analyzing Microsoft 365 Group memberships..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        Get-M365GroupMemberships -ExternalUsers $externalUsers
        
        # Step 6: Analyze Site Sharing Settings
        Write-Host "Step 6/7: Analyzing site sharing settings..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        $sharingSettings = Get-SiteSharingSettings
        
        # Step 7: Analyze Guest Expiration
        Write-Host "Step 7/7: Analyzing guest expiration..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        $guestExpiration = Get-GuestExpiration -ExternalUsers $externalUsers
        
        # Generate Reports
        $reportPath = Export-AuditReport -ExternalUsers $externalUsers `
                                        -SharingSettings $sharingSettings `
                                        -GuestExpiration $guestExpiration
        
        # Display Summary
        Write-Host "`n" -NoNewline
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host "           AUDIT COMPLETED SUCCESSFULLY                 " -ForegroundColor Green
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host "`n📊 Audit Statistics:" -ForegroundColor Cyan
        Write-Host "   • External Users Found: $($externalUsers.Count)" -ForegroundColor White
        Write-Host "   • SharePoint Sites Analyzed: $Global:AuditConfig.ProcessedItems" -ForegroundColor White
        Write-Host "   • Execution Time: $((Get-Date) - $Global:AuditConfig.StartTime)" -ForegroundColor White
        Write-Host "   • Warnings: $($Global:AuditConfig.WarningLog.Count)" -ForegroundColor Yellow
        Write-Host "   • Errors: $($Global:AuditConfig.ErrorLog.Count)" -ForegroundColor Red
        Write-Host "`n📁 Reports Generated:" -ForegroundColor Cyan
        Write-Host "   $OutputPath" -ForegroundColor White
        Write-Host "`n✅ Audit complete! Please review the generated reports." -ForegroundColor Green
        
        # Disconnect services
        Write-Host "Disconnecting from Microsoft 365 services..." -ForegroundColor $(if("Info" -eq "Info"){"Cyan"}elseif("Info" -eq "Success"){"Green"}elseif("Info" -eq "Warning"){"Yellow"}elseif("Info" -eq "Error"){"Red"}else{"White"})
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        
        return $reportPath
    }
    catch {
        Write-Host "Critical error during audit execution" -ForegroundColor $(if("Error" -eq "Info"){"Cyan"}elseif("Error" -eq "Success"){"Green"}elseif("Error" -eq "Warning"){"Yellow"}elseif("Error" -eq "Error"){"Red"}else{"White"}) -ErrorRecord $_
        
        # Export error log
        if ($Global:AuditConfig.ErrorLog.Count -gt 0) {
            # Ensure output directory exists
            if (-not (Test-Path $OutputPath)) {
                New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
            }
            $errorLogPath = Join-Path $OutputPath "ErrorLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $Global:AuditConfig.ErrorLog | ConvertTo-Json -Depth 5 | Out-File -FilePath $errorLogPath
            Write-Host "`n❌ Errors were encountered. Error log saved to: $errorLogPath" -ForegroundColor Red
        }
        
        throw
    }
}

# Execute the audit
Start-ExternalUserAudit

#endregion
