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
    
    Write-Host "🔧 Checking Microsoft Graph module..." -ForegroundColor Cyan
    
    # Check if Microsoft Graph module is already loaded
    $loadedGraphModule = Get-Module -Name "Microsoft.Graph" -ErrorAction SilentlyContinue
    if ($loadedGraphModule) {
        Write-Host "✅ Microsoft Graph module already loaded (version $($loadedGraphModule.Version))" -ForegroundColor Green
        return
    }
    
    # Check for required Microsoft Graph module
    Write-Host "Checking for Microsoft Graph module..." -ForegroundColor Cyan
    
    $installedModule = Get-Module -Name "Microsoft.Graph" -ListAvailable | 
        Where-Object { $_.Version -ge [Version]"1.0.0" } | 
        Sort-Object Version -Descending |
        Select-Object -First 1
    
    if (-not $installedModule) {
        Write-Host "Installing Microsoft Graph module..." -ForegroundColor Yellow
        try {
            $installParams = @{
                Name = "Microsoft.Graph"
                Scope = "CurrentUser"
                Force = $true
                AllowClobber = $true
            }
            
            Install-Module @installParams
            Write-Host "✅ Microsoft Graph module installed successfully" -ForegroundColor Green
            
            # Get the newly installed module
            $installedModule = Get-Module -Name "Microsoft.Graph" -ListAvailable | 
                Sort-Object Version -Descending |
                Select-Object -First 1
        }
        catch {
            Write-Host "❌ Failed to install Microsoft Graph module: $($_.Exception.Message)" -ForegroundColor Red
            throw "Microsoft Graph module is required for this audit script"
        }
    }
    else {
        Write-Host "✅ Microsoft Graph module version $($installedModule.Version) is available" -ForegroundColor Green
    }
    
    # Import Microsoft Graph module if not already loaded
    try {
        Write-Host "Importing Microsoft Graph module..." -ForegroundColor Cyan
        
        Import-Module -Name "Microsoft.Graph" -Force -Global -ErrorAction Stop
        Write-Host "✅ Microsoft Graph module imported successfully" -ForegroundColor Green
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

#endregion

#region External User Discovery

function Get-ExternalUsers {
    [CmdletBinding()]
    param()
    
    $externalUsers = [System.Collections.ArrayList]::new()
    
    try {
        Write-Host "🔍 Discovering external users in tenant..." -ForegroundColor Cyan
        
        # Get Guest Users from Azure AD using Microsoft Graph
        try {
            Write-Host "Retrieving guest users from Azure AD..." -ForegroundColor Cyan
            
            $guestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -ErrorAction Stop
            Write-Host "Successfully retrieved $($guestUsers.Count) guest users" -ForegroundColor Green
            
            foreach ($guest in $guestUsers) {
                try {
                    # Get additional user details
                    $userDetails = Get-MgUser -UserId $guest.Id -Property "Id,UserPrincipalName,DisplayName,Mail,UserType,CompanyName,Department,JobTitle,AccountEnabled,CreatedDateTime,SignInActivity,ExternalUserState,ExternalUserStateChangeDateTime" -ErrorAction SilentlyContinue
                    
                    $userObj = [PSCustomObject]@{
                        Id = $guest.Id
                        Email = if ($userDetails.Mail) { $userDetails.Mail } else { $guest.UserPrincipalName }
                        DisplayName = $userDetails.DisplayName
                        UPN = $userDetails.UserPrincipalName
                        UserType = $userDetails.UserType
                        CompanyName = $userDetails.CompanyName
                        Department = $userDetails.Department
                        JobTitle = $userDetails.JobTitle
                        AccountEnabled = $userDetails.AccountEnabled
                        CreatedDate = if ($userDetails.CreatedDateTime) { [DateTime]$userDetails.CreatedDateTime } else { $null }
                        LastSignIn = if ($userDetails.SignInActivity -and $userDetails.SignInActivity.LastSignInDateTime) { [DateTime]$userDetails.SignInActivity.LastSignInDateTime } else { $null }
                        ExternalUserState = $userDetails.ExternalUserState
                        StateChangeDate = if ($userDetails.ExternalUserStateChangeDateTime) { [DateTime]$userDetails.ExternalUserStateChangeDateTime } else { $null }
                        Source = "Azure AD"
                        # Initialize access collections
                        SharePointSites = [System.Collections.ArrayList]::new()
                        TeamsAccess = [System.Collections.ArrayList]::new()
                        GroupMemberships = [System.Collections.ArrayList]::new()
                        FilesCreated = [System.Collections.ArrayList]::new()
                    }
                    
                    $externalUsers.Add($userObj) | Out-Null
                }
                catch {
                    Write-Host "Warning: Could not get details for user $($guest.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            Write-Host "Found $($guestUsers.Count) guest users in Azure AD" -ForegroundColor Cyan
        }
        catch {
            Write-Host "Error retrieving guest users from Azure AD: $($_.Exception.Message)" -ForegroundColor Yellow
            $guestUsers = @()
        }
        
        Write-Host "Total external users discovered: $($externalUsers.Count)" -ForegroundColor Green
        
        return $externalUsers
    }
    catch {
        Write-Host "Error discovering external users" -ForegroundColor Red -ErrorRecord $_
        throw
    }
}

#endregion

#region SharePoint Site Analysis (Graph API)

function Get-SharePointSiteAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExternalUsers
    )
    
    try {
        Write-Host "📂 Analyzing SharePoint site access via Microsoft Graph..." -ForegroundColor Cyan
        
        # Get all SharePoint sites via Graph API
        try {
            $sites = Get-MgSite -All -ErrorAction Stop
            Write-Host "Found $($sites.Count) SharePoint sites" -ForegroundColor Cyan
        }
        catch {
            Write-Host "Error retrieving SharePoint sites: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "Note: SharePoint sites access requires Sites.Read.All permission" -ForegroundColor Yellow
            return
        }
        
        $Global:AuditConfig.TotalItems = $sites.Count
        $siteIndex = 0
        
        foreach ($site in $sites) {
            $siteIndex++
            $Global:AuditConfig.ProcessedItems = $siteIndex
            $progress = ($siteIndex / $sites.Count) * 100
            
            Write-Progress -Activity "Analyzing SharePoint Sites" `
                          -Status "Processing: $($site.DisplayName)" `
                          -PercentComplete $progress
            
            try {
                # Get site permissions via Graph API
                $sitePermissions = Get-MgSitePermission -SiteId $site.Id -ErrorAction SilentlyContinue
                
                foreach ($permission in $sitePermissions) {
                    # Check if this permission is for an external user
                    if ($permission.GrantedToIdentitiesV2) {
                        foreach ($identity in $permission.GrantedToIdentitiesV2) {
                            $email = $identity.User.Email
                            if ($email) {
                                $externalUser = $ExternalUsers | Where-Object { $_.Email -eq $email }
                                
                                if ($externalUser) {
                                    $siteAccess = [PSCustomObject]@{
                                        SiteId = $site.Id
                                        SiteTitle = $site.DisplayName
                                        SiteUrl = $site.WebUrl
                                        PermissionLevel = $permission.Roles -join ", "
                                        GrantedDate = $permission.CreatedDateTime
                                    }
                                    
                                    $externalUser.SharePointSites.Add($siteAccess) | Out-Null
                                }
                            }
                        }
                    }
                }
            }
            catch {
                Write-Host "Warning: Could not analyze permissions for site $($site.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        Write-Host "✅ SharePoint site analysis complete" -ForegroundColor Green
    }
    catch {
        Write-Host "Error analyzing SharePoint sites" -ForegroundColor Red -ErrorRecord $_
        # Don't throw - continue with rest of audit
    }
}

#endregion

#region Microsoft Teams Analysis (Graph API)

function Get-TeamsAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExternalUsers
    )
    
    try {
        Write-Host "👥 Analyzing Microsoft Teams access via Microsoft Graph..." -ForegroundColor Cyan
        
        # Get all teams using Graph API
        try {
            $teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All -ErrorAction Stop
            Write-Host "Found $($teams.Count) Microsoft Teams" -ForegroundColor Cyan
        }
        catch {
            Write-Host "Error retrieving Teams: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "Trying alternative approach..." -ForegroundColor Yellow
            try {
                # Alternative: Get all groups and filter for teams
                $allGroups = Get-MgGroup -All -ErrorAction Stop
                $teams = $allGroups | Where-Object { $_.ResourceProvisioningOptions -contains "Team" }
                Write-Host "Found $($teams.Count) Microsoft Teams using alternative method" -ForegroundColor Cyan
            }
            catch {
                Write-Host "Error retrieving Teams with alternative method: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "Skipping Teams analysis..." -ForegroundColor Yellow
                return
            }
        }
        
        if ($teams.Count -eq 0) {
            Write-Host "No Teams found or accessible" -ForegroundColor Yellow
            return
        }
        
        $Global:AuditConfig.TotalItems = $teams.Count
        $teamIndex = 0
        
        foreach ($team in $teams) {
            $teamIndex++
            $Global:AuditConfig.ProcessedItems = $teamIndex
            $progress = ($teamIndex / $teams.Count) * 100
            
            Write-Progress -Activity "Analyzing Microsoft Teams" `
                          -Status "Processing: $($team.DisplayName)" `
                          -PercentComplete $progress
            
            try {
                # Get team members
                $members = Get-MgGroupMember -GroupId $team.Id -All -ErrorAction SilentlyContinue
                
                foreach ($member in $members) {
                    try {
                        # Get member details
                        $memberUser = Get-MgUser -UserId $member.Id -Property "Id,UserPrincipalName,Mail,UserType" -ErrorAction SilentlyContinue
                        
                        if ($memberUser -and $memberUser.UserType -eq "Guest") {
                            $email = if ($memberUser.Mail) { $memberUser.Mail } else { $memberUser.UserPrincipalName }
                            
                            $externalUser = $ExternalUsers | Where-Object { $_.Email -eq $email }
                            
                            if ($externalUser) {
                                # Get ownership information
                                $isOwner = $false
                                try {
                                    $owners = Get-MgGroupOwner -GroupId $team.Id -All -ErrorAction SilentlyContinue
                                    $isOwner = $owners | Where-Object { $_.Id -eq $member.Id }
                                }
                                catch {
                                    # Ignore ownership check errors
                                }
                                
                                $teamAccess = [PSCustomObject]@{
                                    TeamId = $team.Id
                                    TeamName = $team.DisplayName
                                    MemberRole = if ($isOwner) { "Owner" } else { "Member" }
                                    JoinedDate = $null # Not available via Graph API
                                }
                                
                                $externalUser.TeamsAccess.Add($teamAccess) | Out-Null
                            }
                        }
                    }
                    catch {
                        Write-Host "Warning: Could not process member for team $($team.DisplayName)" -ForegroundColor Yellow
                    }
                }
            }
            catch {
                Write-Host "Error processing team: $($team.DisplayName)" -ForegroundColor Yellow
            }
        }
        
        Write-Host "✅ Teams analysis complete" -ForegroundColor Green
    }
    catch {
        Write-Host "Error analyzing Microsoft Teams" -ForegroundColor Red -ErrorRecord $_
        # Don't throw - continue with rest of audit
    }
}

#endregion

#region Microsoft 365 Groups Analysis (Graph API)

function Get-M365GroupMemberships {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExternalUsers
    )
    
    try {
        Write-Host "📊 Analyzing Microsoft 365 Group memberships via Microsoft Graph..." -ForegroundColor Cyan
        
        # Get all Microsoft 365 Groups (excluding Teams) using Graph API
        try {
            $groups = Get-MgGroup -Filter "groupTypes/Any(x:x eq 'Unified') and NOT resourceProvisioningOptions/Any(x:x eq 'Team')" -All -ErrorAction Stop
            Write-Host "Found $($groups.Count) Microsoft 365 Groups" -ForegroundColor Cyan
        }
        catch {
            Write-Host "Primary filter method failed, trying alternative approach..." -ForegroundColor Yellow
            try {
                # Alternative: Get all unified groups and filter out teams manually
                $allGroups = Get-MgGroup -Filter "groupTypes/Any(x:x eq 'Unified')" -All -ErrorAction Stop
                $groups = $allGroups | Where-Object { $_.ResourceProvisioningOptions -notcontains "Team" }
                Write-Host "Found $($groups.Count) Microsoft 365 Groups using alternative method" -ForegroundColor Cyan
            }
            catch {
                Write-Host "Error retrieving M365 Groups: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "Skipping M365 Groups analysis..." -ForegroundColor Yellow
                return
            }
        }
        
        if ($groups.Count -eq 0) {
            Write-Host "No M365 Groups found or accessible" -ForegroundColor Yellow
            return
        }
        
        foreach ($group in $groups) {
            try {
                # Get group members
                $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue
                
                foreach ($member in $members) {
                    try {
                        # Get member details
                        $memberUser = Get-MgUser -UserId $member.Id -Property "Id,UserPrincipalName,Mail,UserType" -ErrorAction SilentlyContinue
                        
                        if ($memberUser -and $memberUser.UserType -eq "Guest") {
                            $email = if ($memberUser.Mail) { $memberUser.Mail } else { $memberUser.UserPrincipalName }
                            
                            $externalUser = $ExternalUsers | Where-Object { $_.Email -eq $email }
                            
                            if ($externalUser) {
                                # Get ownership information
                                $isOwner = $false
                                try {
                                    $owners = Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction SilentlyContinue
                                    $isOwner = $owners | Where-Object { $_.Id -eq $member.Id }
                                }
                                catch {
                                    # Ignore ownership check errors
                                }
                                
                                $groupMembership = [PSCustomObject]@{
                                    GroupId = $group.Id
                                    GroupName = $group.DisplayName
                                    MembershipType = if ($isOwner) { "Owner" } else { "Member" }
                                    JoinedDate = $null # Not available via Graph API
                                }
                                
                                $externalUser.GroupMemberships.Add($groupMembership) | Out-Null
                            }
                        }
                    }
                    catch {
                        Write-Host "Warning: Could not process member for group $($group.DisplayName)" -ForegroundColor Yellow
                    }
                }
            }
            catch {
                Write-Host "Error processing group: $($group.DisplayName)" -ForegroundColor Yellow
            }
        }
        
        Write-Host "✅ Microsoft 365 Groups analysis complete" -ForegroundColor Green
    }
    catch {
        Write-Host "Error analyzing Microsoft 365 Groups" -ForegroundColor Red -ErrorRecord $_
        # Don't throw - continue with rest of audit
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
        Write-Host "⏰ Analyzing guest expiration settings..." -ForegroundColor Cyan
        
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
                DaysSinceCreation = if ($user.CreatedDate) { ((Get-Date) - $user.CreatedDate).Days } else { "Unknown" }
                DaysSinceLastSignIn = if ($user.LastSignIn) { ((Get-Date) - $user.LastSignIn).Days } else { "Never signed in" }
                ExpirationStatus = "Active"
                RecommendedAction = ""
            }
            
            # Determine expiration status and recommendations
            if ($user.LastSignIn -and (((Get-Date) - $user.LastSignIn).Days -gt 90)) {
                $expirationObj.ExpirationStatus = "Inactive"
                $expirationObj.RecommendedAction = "Consider removing - inactive for over 90 days"
            }
            elseif ($user.LastSignIn -and (((Get-Date) - $user.LastSignIn).Days -gt 60)) {
                $expirationObj.ExpirationStatus = "Warning"
                $expirationObj.RecommendedAction = "Review access - inactive for over 60 days"
            }
            elseif (-not $user.LastSignIn) {
                $expirationObj.ExpirationStatus = "Never Signed In"
                $expirationObj.RecommendedAction = "Review - user has never signed in"
            }
            
            if (-not $user.AccountEnabled) {
                $expirationObj.ExpirationStatus = "Disabled"
                $expirationObj.RecommendedAction = "Account disabled - review for removal"
            }
            
            $guestExpirationReport.Add($expirationObj) | Out-Null
        }
        
        Write-Host "✅ Guest expiration analysis complete" -ForegroundColor Green
        
        return $guestExpirationReport
    }
    catch {
        Write-Host "Error analyzing guest expiration" -ForegroundColor Red -ErrorRecord $_
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
        [object[]]$GuestExpiration
    )
    
    try {
        Write-Host "📄 Generating audit reports..." -ForegroundColor Cyan
        
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
                
                # Risk Assessment
                RiskLevel = Get-UserRiskLevel -User $user
            }
            
            $mainReport.Add($reportEntry) | Out-Null
        }
        
        # Export main report
        $mainReportPath = Join-Path $OutputPath "ExternalUserAudit_GraphOnly_$timestamp.csv"
        $mainReport | Export-Csv -Path $mainReportPath -NoTypeInformation
        Write-Host "✅ Main report exported: $mainReportPath" -ForegroundColor Green
        
        # Export individual detailed reports if requested
        if ($ExportIndividualReports) {
            # SharePoint Access Details
            $spReport = [System.Collections.ArrayList]::new()
            foreach ($user in $ExternalUsers) {
                foreach ($site in $user.SharePointSites) {
                    $spEntry = [PSCustomObject]@{
                        UserEmail = $user.Email
                        UserDisplayName = $user.DisplayName
                        SiteTitle = $site.SiteTitle
                        SiteUrl = $site.SiteUrl
                        PermissionLevel = $site.PermissionLevel
                        GrantedDate = $site.GrantedDate
                    }
                    $spReport.Add($spEntry) | Out-Null
                }
            }
            
            if ($spReport.Count -gt 0) {
                $spReportPath = Join-Path $OutputPath "SharePointAccess_$timestamp.csv"
                $spReport | Export-Csv -Path $spReportPath -NoTypeInformation
                Write-Host "✅ SharePoint access report exported: $spReportPath" -ForegroundColor Green
            }
            
            # Teams Access Details
            $teamsReport = [System.Collections.ArrayList]::new()
            foreach ($user in $ExternalUsers) {
                foreach ($team in $user.TeamsAccess) {
                    $teamEntry = [PSCustomObject]@{
                        UserEmail = $user.Email
                        UserDisplayName = $user.DisplayName
                        TeamName = $team.TeamName
                        MemberRole = $team.MemberRole
                        JoinedDate = $team.JoinedDate
                    }
                    $teamsReport.Add($teamEntry) | Out-Null
                }
            }
            
            if ($teamsReport.Count -gt 0) {
                $teamsReportPath = Join-Path $OutputPath "TeamsAccess_$timestamp.csv"
                $teamsReport | Export-Csv -Path $teamsReportPath -NoTypeInformation
                Write-Host "✅ Teams access report exported: $teamsReportPath" -ForegroundColor Green
            }
        }
        
        # Export Guest Expiration Report
        $expirationReportPath = Join-Path $OutputPath "GuestExpiration_$timestamp.csv"
        $GuestExpiration | Export-Csv -Path $expirationReportPath -NoTypeInformation
        Write-Host "✅ Guest expiration report exported: $expirationReportPath" -ForegroundColor Green
        
        # Generate Executive Summary
        Export-ExecutiveSummary -ExternalUsers $ExternalUsers -OutputPath $OutputPath -Timestamp $timestamp
        
        Write-Host "✅ All reports generated successfully" -ForegroundColor Green
        
        return $mainReportPath
    }
    catch {
        Write-Host "Error generating reports" -ForegroundColor Red -ErrorRecord $_
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
    if ($User.SharePointSites | Where-Object { $_.PermissionLevel -like "*Owner*" -or $_.PermissionLevel -like "*Administrator*" }) { $riskScore += 2 }
    if ($User.TeamsAccess | Where-Object { $_.MemberRole -eq "Owner" }) { $riskScore += 2 }
    if ($User.GroupMemberships | Where-Object { $_.MembershipType -eq "Owner" }) { $riskScore += 1 }
    
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
# Microsoft 365 External User Audit - Executive Summary (Graph API Only)
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

## 📊 Overview Statistics
- Total External Users: $($ExternalUsers.Count)
- Active Users (signed in last 30 days): $(($ExternalUsers | Where-Object { $_.LastSignIn -and ((Get-Date) - $_.LastSignIn).Days -le 30 }).Count)
- Inactive Users (>90 days): $(($ExternalUsers | Where-Object { $_.LastSignIn -and ((Get-Date) - $_.LastSignIn).Days -gt 90 }).Count)
- Never Signed In: $(($ExternalUsers | Where-Object { -not $_.LastSignIn }).Count)
- Disabled Accounts: $(($ExternalUsers | Where-Object { -not $_.AccountEnabled }).Count)

## 🔐 Access Distribution
- Users with SharePoint Access: $(($ExternalUsers | Where-Object { $_.SharePointSites.Count -gt 0 }).Count)
- Users with Teams Access: $(($ExternalUsers | Where-Object { $_.TeamsAccess.Count -gt 0 }).Count)
- Users with Group Memberships: $(($ExternalUsers | Where-Object { $_.GroupMemberships.Count -gt 0 }).Count)

## ⚠️ Risk Assessment
- High Risk Users: $(($ExternalUsers | Where-Object { (Get-UserRiskLevel -User $_) -eq "High" }).Count)
- Medium Risk Users: $(($ExternalUsers | Where-Object { (Get-UserRiskLevel -User $_) -eq "Medium" }).Count)
- Low Risk Users: $(($ExternalUsers | Where-Object { (Get-UserRiskLevel -User $_) -eq "Low" }).Count)

## 🎯 Recommendations
1. Review and remove access for inactive users (>90 days without sign-in)
2. Audit high-risk users with extensive permissions
3. Implement regular access reviews for external users
4. Consider implementing guest expiration policies
5. Review users who have never signed in - they may be unused invitations

## 📝 Audit Details
- Audit Duration: $((Get-Date) - $Global:AuditConfig.StartTime)
- Sites Analyzed: $Global:AuditConfig.ProcessedItems
- Warnings Encountered: $($Global:AuditConfig.WarningLog.Count)
- Errors Encountered: $($Global:AuditConfig.ErrorLog.Count)
- API Used: Microsoft Graph API Only (no SharePoint PnP or Exchange Online)

## 🔧 Technical Notes
This audit was performed using Microsoft Graph API exclusively, providing:
- Comprehensive user information from Azure AD
- Teams and Groups membership analysis
- SharePoint site permissions (where available)
- Simplified authentication and improved reliability
"@
    
    $summaryPath = Join-Path $OutputPath "ExecutiveSummary_$Timestamp.md"
    $summary | Out-File -FilePath $summaryPath -Encoding UTF8
    
    Write-Host "✅ Executive summary generated: $summaryPath" -ForegroundColor Green
}

#endregion

#region Main Execution

function Start-ExternalUserAudit {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host @"
╔══════════════════════════════════════════════════════════════╗
║   Microsoft 365 External User Audit Tool - Graph API Only   ║
║                     Version 3.0.0                            ║
╚══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan
        
        Write-Host "Starting Microsoft 365 External User Audit (Microsoft Graph Only)..." -ForegroundColor Cyan
        Write-Host "Output Path: $OutputPath" -ForegroundColor Cyan
        
        # Step 0: Validate and install required modules
        Write-Host "Step 0/5: Checking Microsoft Graph module..." -ForegroundColor Cyan
        Install-RequiredModules
        
        # Step 1: Authentication
        Write-Host "Step 1/5: Authenticating to Microsoft Graph..." -ForegroundColor Cyan
        Connect-M365Services
        
        # Step 2: Discover External Users
        Write-Host "Step 2/5: Discovering external users..." -ForegroundColor Cyan
        $externalUsers = Get-ExternalUsers
        
        if ($externalUsers.Count -eq 0) {
            Write-Host "No external users found in the tenant" -ForegroundColor Yellow
            return
        }
        
        # Step 3: Analyze SharePoint Access
        Write-Host "Step 3/5: Analyzing SharePoint site access..." -ForegroundColor Cyan
        Get-SharePointSiteAccess -ExternalUsers $externalUsers
        
        # Step 4: Analyze Teams and Groups Access
        Write-Host "Step 4/5: Analyzing Microsoft Teams and Groups access..." -ForegroundColor Cyan
        Get-TeamsAccess -ExternalUsers $externalUsers
        Get-M365GroupMemberships -ExternalUsers $externalUsers
        
        # Step 5: Analyze Guest Expiration
        Write-Host "Step 5/5: Analyzing guest expiration..." -ForegroundColor Cyan
        $guestExpiration = Get-GuestExpiration -ExternalUsers $externalUsers
        
        # Generate Reports
        $reportPath = Export-AuditReport -ExternalUsers $externalUsers -GuestExpiration $guestExpiration
        
        # Display Summary
        Write-Host "`n" -NoNewline
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host "           AUDIT COMPLETED SUCCESSFULLY                 " -ForegroundColor Green
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host "`n📊 Audit Statistics:" -ForegroundColor Cyan
        Write-Host "   • External Users Found: $($externalUsers.Count)" -ForegroundColor White
        Write-Host "   • Sites Analyzed: $Global:AuditConfig.ProcessedItems" -ForegroundColor White
        Write-Host "   • Execution Time: $((Get-Date) - $Global:AuditConfig.StartTime)" -ForegroundColor White
        Write-Host "   • Warnings: $($Global:AuditConfig.WarningLog.Count)" -ForegroundColor Yellow
        Write-Host "   • Errors: $($Global:AuditConfig.ErrorLog.Count)" -ForegroundColor Red
        Write-Host "`n📁 Reports Generated:" -ForegroundColor Cyan
        Write-Host "   $OutputPath" -ForegroundColor White
        Write-Host "`n✅ Audit complete! Using Microsoft Graph API only for improved reliability." -ForegroundColor Green
        
        # Disconnect services
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        
        return $reportPath
    }
    catch {
        Write-Host "Critical error during audit execution" -ForegroundColor Red -ErrorRecord $_
        
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
