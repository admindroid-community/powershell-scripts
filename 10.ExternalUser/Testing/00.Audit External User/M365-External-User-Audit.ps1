#Requires -Version 7.0
#Requires -Modules Microsoft.Graph, PnP.PowerShell, ExchangeOnlineManagement

<#
.SYNOPSIS
    Comprehensive Microsoft 365 External User Audit Script
.DESCRIPTION
    Enterprise-grade audit solution for external users across SharePoint Online, Microsoft Teams, and Microsoft 365 Groups
.AUTHOR
    Microsoft 365 Security SME
.VERSION
    2.0.0
.NOTES
    Implements zero-trust security model with comprehensive error handling and compliance reporting
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$TenantId = "",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\AuditReports",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeFileAnalysis,
    
    [Parameter(Mandatory = $false)]
    [switch]$DetailedPermissions,
    
    [Parameter(Mandatory = $false)]
    [int]$DaysToAudit = 90,
    
    [Parameter(Mandatory = $false)]
    [string[]]$SpecificSites = @(),
    
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
    )
    SPOScopes = @(
        "Sites.FullControl.All"
        "User.Read.All"
        "Group.ReadWrite.All"
    )
}

# 📝 Logging Configuration
function Write-AuditLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Debug')]
        [string]$Level = 'Info',
        
        [Parameter(Mandatory = $false)]
        [object]$ErrorRecord
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    
    switch ($Level) {
        'Info'    { Write-Host $logEntry -ForegroundColor Cyan }
        'Warning' { 
            Write-Warning $logEntry
            $Global:AuditConfig.WarningLog.Add($logEntry) | Out-Null
        }
        'Error'   { 
            Write-Error $logEntry
            if ($ErrorRecord) {
                $Global:AuditConfig.ErrorLog.Add(@{
                    Message = $logEntry
                    Error = $ErrorRecord
                    Stack = $ErrorRecord.ScriptStackTrace
                }) | Out-Null
            }
        }
        'Success' { Write-Host $logEntry -ForegroundColor Green }
        'Debug'   { Write-Debug $logEntry }
    }
    
    # Append to log file
    $logFile = Join-Path $OutputPath "AuditLog_$(Get-Date -Format 'yyyyMMdd').log"
    Add-Content -Path $logFile -Value $logEntry -ErrorAction SilentlyContinue
}

#endregion

#region Authentication Module

function Connect-M365Services {
    [CmdletBinding()]
    param()
    
    try {
        Write-AuditLog "🔐 Initiating Microsoft 365 authentication sequence..." -Level Info
        
        # Connect to Microsoft Graph
        Write-AuditLog "Connecting to Microsoft Graph..." -Level Info
        
        $graphParams = @{
            Scopes = $Global:AuditConfig.GraphScopes
            NoWelcome = $true
        }
        
        if ($TenantId) {
            $graphParams.Add("TenantId", $TenantId)
        }
        
        try {
            Connect-MgGraph @graphParams -ErrorAction Stop
            
            # Wait a moment for the connection to stabilize
            Start-Sleep -Seconds 2
            
            $context = Get-MgContext -ErrorAction Stop
            Write-AuditLog "✅ Connected to Microsoft Graph as: $($context.Account)" -Level Success
        }
        catch {
            Write-AuditLog "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level Error
            throw
        }
        
        # Connect to SharePoint Online
        Write-AuditLog "Connecting to SharePoint Online..." -Level Info
        
        try {
            if (-not $TenantId) {
                # Try alternative method to get tenant info
                $context = Get-MgContext -ErrorAction Stop
                $TenantId = $context.TenantId
            }
            
            # Use hardcoded admin URL pattern as fallback
            $adminUrl = "https://m365x22747677-admin.sharepoint.com"
            Write-AuditLog "Using SharePoint Admin URL: $adminUrl" -Level Info
        }
        catch {
            Write-AuditLog "Could not determine tenant info, using default admin URL pattern" -Level Warning
            $adminUrl = "https://m365x22747677-admin.sharepoint.com"
        }
        
        Connect-PnPOnline -Url $adminUrl -ClientId "afe1b358-534b-4c96-abb9-ecea5d5f2e5d" -Interactive -ErrorAction Stop
        
        Write-AuditLog "✅ Connected to SharePoint Online: $adminUrl" -Level Success
        
        # Connect to Exchange Online (for unified group management)
        Write-AuditLog "Connecting to Exchange Online..." -Level Info
        try {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Write-AuditLog "✅ Connected to Exchange Online successfully" -Level Success
        }
        catch {
            Write-AuditLog "Warning: Could not connect to Exchange Online. Continuing with limited functionality..." -Level Warning
            Write-AuditLog "Exchange Online error: $($_.Exception.Message)" -Level Warning
        }
        
        Write-AuditLog "✅ All services connected successfully" -Level Success
        
        return $true
    }
    catch {
        Write-AuditLog "Failed to authenticate to Microsoft 365 services" -Level Error -ErrorRecord $_
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
        Write-AuditLog "🔍 Discovering external users in tenant..." -Level Info
        
        # Method 1: Get Guest Users from Azure AD
        $guestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -Property @(
            'Id', 'DisplayName', 'Mail', 'UserPrincipalName', 'CreatedDateTime',
            'ExternalUserState', 'ExternalUserStateChangeDateTime', 'CompanyName',
            'Department', 'JobTitle', 'AccountEnabled', 'SignInActivity'
        )
        
        foreach ($guest in $guestUsers) {
            $userObj = [PSCustomObject]@{
                Id = $guest.Id
                DisplayName = $guest.DisplayName
                Email = $guest.Mail ?? $guest.UserPrincipalName
                UPN = $guest.UserPrincipalName
                UserType = "Guest"
                CreatedDate = $guest.CreatedDateTime
                ExternalUserState = $guest.ExternalUserState
                StateChangeDate = $guest.ExternalUserStateChangeDateTime
                CompanyName = $guest.CompanyName
                Department = $guest.Department
                JobTitle = $guest.JobTitle
                AccountEnabled = $guest.AccountEnabled
                LastSignIn = $guest.SignInActivity.LastSignInDateTime
                Source = "Azure AD"
                SharePointSites = [System.Collections.ArrayList]::new()
                TeamsAccess = [System.Collections.ArrayList]::new()
                GroupMemberships = [System.Collections.ArrayList]::new()
                FilesCreated = [System.Collections.ArrayList]::new()
                Permissions = [System.Collections.ArrayList]::new()
            }
            
            $externalUsers.Add($userObj) | Out-Null
        }
        
        Write-AuditLog "Found $($guestUsers.Count) guest users in Azure AD" -Level Info
        
        # Method 2: Get External Users from SharePoint
        $spoExternalUsers = Get-PnPExternalUser -PageSize 50 -Position 0
        
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
        
        Write-AuditLog "Total external users discovered: $($externalUsers.Count)" -Level Success
        
        return $externalUsers
    }
    catch {
        Write-AuditLog "Error discovering external users" -Level Error -ErrorRecord $_
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
        Write-AuditLog "📂 Analyzing SharePoint site access..." -Level Info
        
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
                # Connect to specific site
                Connect-PnPOnline -Url $site.Url -ClientId "afe1b358-534b-4c96-abb9-ecea5d5f2e5d" -Interactive -ErrorAction Stop
                
                # Get site users
                $siteUsers = Get-PnPUser -WithRightsAssigned
                
                foreach ($siteUser in $siteUsers) {
                    if ($siteUser.LoginName -match "#ext#" -or $siteUser.UserPrincipalName -match "#EXT#") {
                        $extUser = $ExternalUsers | Where-Object { 
                            $_.Email -eq $siteUser.Email -or 
                            $_.UPN -eq $siteUser.UserPrincipalName 
                        }
                        
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
                            
                            # Get detailed permissions if requested
                            if ($DetailedPermissions) {
                                $permissions = Get-UserSitePermissions -SiteUrl $site.Url -UserEmail $extUser.Email
                                $extUser.Permissions.Add($permissions) | Out-Null
                            }
                        }
                    }
                }
                
                # Analyze files created by external users if requested
                if ($IncludeFileAnalysis) {
                    $files = Get-ExternalUserFiles -SiteUrl $site.Url -ExternalUsers $ExternalUsers
                    
                    foreach ($file in $files) {
                        $creator = $ExternalUsers | Where-Object { $_.Email -eq $file.CreatedBy }
                        if ($creator) {
                            $creator.FilesCreated.Add($file) | Out-Null
                        }
                    }
                }
            }
            catch {
                Write-AuditLog "Error processing site: $($site.Url)" -Level Warning
            }
        }
        
        Write-AuditLog "✅ SharePoint site analysis complete" -Level Success
    }
    catch {
        Write-AuditLog "Error analyzing SharePoint sites" -Level Error -ErrorRecord $_
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
        Write-AuditLog "Error getting detailed permissions for $UserEmail" -Level Warning
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
                Write-AuditLog "Error processing library: $($list.Title)" -Level Warning
            }
        }
    }
    catch {
        Write-AuditLog "Error retrieving files for external users" -Level Warning
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
        Write-AuditLog "👥 Analyzing Microsoft Teams access..." -Level Info
        
        # Get all teams
        $teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All
        
        $Global:AuditConfig.TotalItems = $teams.Count
        $Global:AuditConfig.ProcessedItems = 0
        
        foreach ($team in $teams) {
            $Global:AuditConfig.ProcessedItems++
            $progress = ($Global:AuditConfig.ProcessedItems / $Global:AuditConfig.TotalItems) * 100
            
            Write-Progress -Activity "Analyzing Microsoft Teams" `
                          -Status "Processing: $($team.DisplayName)" `
                          -PercentComplete $progress
            
            try {
                # Get team members
                $members = Get-MgGroupMember -GroupId $team.Id -All
                
                foreach ($member in $members) {
                    $memberDetails = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
                    
                    if ($memberDetails.UserType -eq "Guest") {
                        $extUser = $ExternalUsers | Where-Object { $_.Id -eq $memberDetails.Id }
                        
                        if ($extUser) {
                            # Get member role
                            $owners = Get-MgGroupOwner -GroupId $team.Id -All
                            $isOwner = $owners.Id -contains $member.Id
                            
                            $teamAccess = [PSCustomObject]@{
                                TeamId = $team.Id
                                TeamName = $team.DisplayName
                                TeamDescription = $team.Description
                                TeamVisibility = $team.Visibility
                                MemberRole = $isOwner ? "Owner" : "Member"
                                JoinedDate = $null
                                TeamArchived = $team.IsArchived
                                TeamSiteUrl = $team.SharePointSiteUrl
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
            }
            catch {
                Write-AuditLog "Error processing team: $($team.DisplayName)" -Level Warning
            }
        }
        
        Write-AuditLog "✅ Teams analysis complete" -Level Success
    }
    catch {
        Write-AuditLog "Error analyzing Microsoft Teams" -Level Error -ErrorRecord $_
        throw
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
        Write-AuditLog "📊 Analyzing Microsoft 365 Group memberships..." -Level Info
        
        # Get all Microsoft 365 Groups (excluding Teams)
        $groups = Get-MgGroup -Filter "groupTypes/Any(x:x eq 'Unified') and NOT resourceProvisioningOptions/Any(x:x eq 'Team')" -All
        
        foreach ($group in $groups) {
            try {
                $members = Get-MgGroupMember -GroupId $group.Id -All
                
                foreach ($member in $members) {
                    $memberDetails = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
                    
                    if ($memberDetails.UserType -eq "Guest") {
                        $extUser = $ExternalUsers | Where-Object { $_.Id -eq $memberDetails.Id }
                        
                        if ($extUser) {
                            $owners = Get-MgGroupOwner -GroupId $group.Id -All
                            $isOwner = $owners.Id -contains $member.Id
                            
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
            }
            catch {
                Write-AuditLog "Error processing group: $($group.DisplayName)" -Level Warning
            }
        }
        
        Write-AuditLog "✅ Microsoft 365 Groups analysis complete" -Level Success
    }
    catch {
        Write-AuditLog "Error analyzing Microsoft 365 Groups" -Level Error -ErrorRecord $_
        throw
    }
}

#endregion

#region Site Sharing Settings Analysis

function Get-SiteSharingSettings {
    [CmdletBinding()]
    param()
    
    $sharingSettings = [System.Collections.ArrayList]::new()
    
    try {
        Write-AuditLog "🔗 Analyzing site sharing settings..." -Level Info
        
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
        
        Write-AuditLog "✅ Site sharing settings analysis complete" -Level Success
        
        return $sharingSettings
    }
    catch {
        Write-AuditLog "Error analyzing site sharing settings" -Level Error -ErrorRecord $_
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
        Write-AuditLog "⏰ Analyzing guest expiration settings..." -Level Info
        
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
        
        Write-AuditLog "✅ Guest expiration analysis complete" -Level Success
        
        return $guestExpirationReport
    }
    catch {
        Write-AuditLog "Error analyzing guest expiration" -Level Error -ErrorRecord $_
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
        Write-AuditLog "📄 Generating audit reports..." -Level Info
        
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
        Write-AuditLog "✅ Main report exported: $mainReportPath" -Level Success
        
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
                Write-AuditLog "✅ SharePoint report exported: $spReportPath" -Level Success
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
                Write-AuditLog "✅ Teams report exported: $teamsReportPath" -Level Success
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
                    Write-AuditLog "✅ Files report exported: $filesReportPath" -Level Success
                }
            }
        }
        
        # Export Site Sharing Settings
        $sharingReportPath = Join-Path $OutputPath "SiteSharingSettings_$timestamp.csv"
        $SharingSettings | Export-Csv -Path $sharingReportPath -NoTypeInformation
        Write-AuditLog "✅ Sharing settings report exported: $sharingReportPath" -Level Success
        
        # Export Guest Expiration Report
        $expirationReportPath = Join-Path $OutputPath "GuestExpiration_$timestamp.csv"
        $GuestExpiration | Export-Csv -Path $expirationReportPath -NoTypeInformation
        Write-AuditLog "✅ Guest expiration report exported: $expirationReportPath" -Level Success
        
        # Generate Executive Summary
        Export-ExecutiveSummary -ExternalUsers $ExternalUsers -OutputPath $OutputPath -Timestamp $timestamp
        
        Write-AuditLog "✅ All reports generated successfully" -Level Success
        
        return $mainReportPath
    }
    catch {
        Write-AuditLog "Error generating reports" -Level Error -ErrorRecord $_
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
    
    Write-AuditLog "✅ Executive summary generated: $summaryPath" -Level Success
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
        
        Write-AuditLog "Starting Microsoft 365 External User Audit..." -Level Info
        Write-AuditLog "Output Path: $OutputPath" -Level Info
        
        # Step 1: Authentication
        Write-AuditLog "Step 1/7: Authenticating to Microsoft 365..." -Level Info
        Connect-M365Services
        
        # Step 2: Discover External Users
        Write-AuditLog "Step 2/7: Discovering external users..." -Level Info
        $externalUsers = Get-ExternalUsers
        
        if ($externalUsers.Count -eq 0) {
            Write-AuditLog "No external users found in the tenant" -Level Warning
            return
        }
        
        # Step 3: Analyze SharePoint Access
        Write-AuditLog "Step 3/7: Analyzing SharePoint site access..." -Level Info
        Get-SharePointSiteAccess -ExternalUsers $externalUsers
        
        # Step 4: Analyze Teams Access
        Write-AuditLog "Step 4/7: Analyzing Microsoft Teams access..." -Level Info
        Get-TeamsAccess -ExternalUsers $externalUsers
        
        # Step 5: Analyze Group Memberships
        Write-AuditLog "Step 5/7: Analyzing Microsoft 365 Group memberships..." -Level Info
        Get-M365GroupMemberships -ExternalUsers $externalUsers
        
        # Step 6: Analyze Site Sharing Settings
        Write-AuditLog "Step 6/7: Analyzing site sharing settings..." -Level Info
        $sharingSettings = Get-SiteSharingSettings
        
        # Step 7: Analyze Guest Expiration
        Write-AuditLog "Step 7/7: Analyzing guest expiration..." -Level Info
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
        Write-AuditLog "Disconnecting from Microsoft 365 services..." -Level Info
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        
        return $reportPath
    }
    catch {
        Write-AuditLog "Critical error during audit execution" -Level Error -ErrorRecord $_
        
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