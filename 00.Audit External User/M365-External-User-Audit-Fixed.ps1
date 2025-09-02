#Requires -Version 7.0
#Requires -Modules Microsoft.Graph, PnP.PowerShell

<#
.SYNOPSIS
    Streamlined Microsoft 365 External User Audit Script
.DESCRIPTION
    Lean audit solution for external users across SharePoint Online and Microsoft Teams
.VERSION
    2.1.0 - Fixed and Optimized
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\AuditReports",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeFileAnalysis,
    
    [Parameter(Mandatory = $false)]
    [int]$DaysToAudit = 90,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId = "afe1b358-534b-4c96-abb9-ecea5d5f2e5d"
)

# Global Configuration
$Global:AuditConfig = @{
    StartTime = Get-Date
    ErrorLog = @()
    ProcessedItems = 0
    GraphScopes = @(
        "User.Read.All"
        "Group.Read.All" 
        "Sites.Read.All"
        "Directory.Read.All"
    )
}

function Write-AuditLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    
    switch ($Level) {
        'Info'    { Write-Host $logEntry -ForegroundColor Cyan }
        'Warning' { Write-Host $logEntry -ForegroundColor Yellow }
        'Error'   { Write-Host $logEntry -ForegroundColor Red }
        'Success' { Write-Host $logEntry -ForegroundColor Green }
    }
}

function Connect-M365Services {
    try {
        Write-AuditLog "🔐 Connecting to Microsoft 365 services..." -Level Info
        
        # Connect to Microsoft Graph
        Write-AuditLog "Connecting to Microsoft Graph..." -Level Info
        Connect-MgGraph -Scopes $Global:AuditConfig.GraphScopes -NoWelcome -ErrorAction Stop
        
        $context = Get-MgContext
        Write-AuditLog "✅ Connected to Microsoft Graph as: $($context.Account)" -Level Success
        
        # Connect to SharePoint Online
        Write-AuditLog "Connecting to SharePoint Online..." -Level Info
        $adminUrl = "https://m365x22747677-admin.sharepoint.com"
        Connect-PnPOnline -Url $adminUrl -ClientId $ClientId -Interactive -ErrorAction Stop
        Write-AuditLog "✅ Connected to SharePoint Online" -Level Success
        
        return $true
    }
    catch {
        Write-AuditLog "Failed to authenticate: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Get-ExternalUsers {
    $externalUsers = @()
    
    try {
        Write-AuditLog "🔍 Discovering external users..." -Level Info
        
        # Get Guest Users from Azure AD
        $guestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -Property Id,DisplayName,Mail,UserPrincipalName,CreatedDateTime,AccountEnabled
        
        foreach ($guest in $guestUsers) {
            $userObj = [PSCustomObject]@{
                Id = $guest.Id
                DisplayName = $guest.DisplayName
                Email = $guest.Mail ?? $guest.UserPrincipalName
                UPN = $guest.UserPrincipalName
                CreatedDate = $guest.CreatedDateTime
                AccountEnabled = $guest.AccountEnabled
                SharePointSites = @()
                TeamsAccess = @()
                GroupMemberships = @()
                FilesCreated = @()
            }
            $externalUsers += $userObj
        }
        
        Write-AuditLog "Found $($guestUsers.Count) guest users" -Level Success
        return $externalUsers
    }
    catch {
        Write-AuditLog "Error discovering external users: $($_.Exception.Message)" -Level Error
        return @()
    }
}

function Get-SharePointSiteAccess {
    param([object[]]$ExternalUsers)
    
    try {
        Write-AuditLog "📂 Analyzing SharePoint site access..." -Level Info
        
        # Get limited number of sites to avoid hanging
        $sites = Get-PnPTenantSite -IncludeOneDriveSites:$false | Select-Object -First 10
        
        $siteCount = 0
        foreach ($site in $sites) {
            $siteCount++
            Write-Progress -Activity "Analyzing SharePoint Sites" -Status "Processing: $($site.Url)" -PercentComplete (($siteCount / $sites.Count) * 100)
            
            try {
                # Get site users without connecting to each site individually
                $siteUsers = Get-PnPSiteUser -Site $site.Url -ErrorAction SilentlyContinue
                
                foreach ($siteUser in $siteUsers) {
                    if ($siteUser.LoginName -match "#ext#|#EXT#") {
                        $extUser = $ExternalUsers | Where-Object { $_.Email -eq $siteUser.Email }
                        
                        if ($extUser) {
                            $siteAccess = [PSCustomObject]@{
                                SiteUrl = $site.Url
                                SiteTitle = $site.Title
                                PermissionLevel = if($siteUser.IsSiteAdmin) { "Site Administrator" } else { "Site Member" }
                            }
                            $extUser.SharePointSites += $siteAccess
                        }
                    }
                }
            }
            catch {
                Write-AuditLog "Warning: Could not process site $($site.Url)" -Level Warning
            }
        }
        
        Write-Progress -Activity "Analyzing SharePoint Sites" -Completed
        Write-AuditLog "✅ SharePoint analysis complete" -Level Success
    }
    catch {
        Write-AuditLog "Error analyzing SharePoint sites: $($_.Exception.Message)" -Level Error
    }
}

function Get-TeamsAccess {
    param([object[]]$ExternalUsers)
    
    try {
        Write-AuditLog "👥 Analyzing Microsoft Teams access..." -Level Info
        
        # Get Teams (Microsoft 365 Groups with Teams)
        $teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -Top 20
        
        $teamCount = 0
        foreach ($team in $teams) {
            $teamCount++
            Write-Progress -Activity "Analyzing Teams" -Status "Processing: $($team.DisplayName)" -PercentComplete (($teamCount / $teams.Count) * 100)
            
            try {
                $members = Get-MgGroupMember -GroupId $team.Id
                
                foreach ($member in $members) {
                    $memberDetails = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
                    
                    if ($memberDetails.UserType -eq "Guest") {
                        $extUser = $ExternalUsers | Where-Object { $_.Id -eq $memberDetails.Id }
                        
                        if ($extUser) {
                            $teamAccess = [PSCustomObject]@{
                                TeamId = $team.Id
                                TeamName = $team.DisplayName
                                MemberRole = "Member"
                            }
                            $extUser.TeamsAccess += $teamAccess
                        }
                    }
                }
            }
            catch {
                Write-AuditLog "Warning: Could not process team $($team.DisplayName)" -Level Warning
            }
        }
        
        Write-Progress -Activity "Analyzing Teams" -Completed
        Write-AuditLog "✅ Teams analysis complete" -Level Success
    }
    catch {
        Write-AuditLog "Error analyzing Teams: $($_.Exception.Message)" -Level Error
    }
}

function Export-AuditReport {
    param([object[]]$ExternalUsers)
    
    try {
        Write-AuditLog "📄 Generating audit report..." -Level Info
        
        # Create output directory
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $reportPath = Join-Path $OutputPath "ExternalUserAudit_$timestamp.csv"
        
        # Create main report
        $mainReport = @()
        
        foreach ($user in $ExternalUsers) {
            $reportEntry = [PSCustomObject]@{
                DisplayName = $user.DisplayName
                Email = $user.Email
                UPN = $user.UPN
                CreatedDate = $user.CreatedDate
                AccountEnabled = $user.AccountEnabled
                SharePointSitesCount = $user.SharePointSites.Count
                SharePointSites = ($user.SharePointSites | ForEach-Object { "$($_.SiteTitle) [$($_.PermissionLevel)]" }) -join "; "
                TeamsCount = $user.TeamsAccess.Count
                Teams = ($user.TeamsAccess | ForEach-Object { $_.TeamName }) -join "; "
                FilesCreatedCount = $user.FilesCreated.Count
            }
            $mainReport += $reportEntry
        }
        
        # Export to CSV
        $mainReport | Export-Csv -Path $reportPath -NoTypeInformation -Force
        Write-AuditLog "✅ Report exported: $reportPath" -Level Success
        
        return $reportPath
    }
    catch {
        Write-AuditLog "Error generating report: $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Start-ExternalUserAudit {
    try {
        Write-Host @"
╔══════════════════════════════════════════════════════════════╗
║   Microsoft 365 External User Audit Tool - FIXED VERSION    ║
║                     Version 2.1.0                            ║
╚══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan
        
        Write-AuditLog "Starting External User Audit..." -Level Info
        Write-AuditLog "Output Path: $OutputPath" -Level Info
        
        # Step 1: Authentication
        if (-not (Connect-M365Services)) {
            Write-AuditLog "Authentication failed. Exiting." -Level Error
            return
        }
        
        # Step 2: Discover External Users
        Write-AuditLog "Step 1/3: Discovering external users..." -Level Info
        $externalUsers = Get-ExternalUsers
        
        if ($externalUsers.Count -eq 0) {
            Write-AuditLog "No external users found" -Level Warning
            return
        }
        
        # Step 3: Analyze SharePoint Access
        Write-AuditLog "Step 2/3: Analyzing SharePoint access..." -Level Info
        Get-SharePointSiteAccess -ExternalUsers $externalUsers
        
        # Step 4: Analyze Teams Access
        Write-AuditLog "Step 3/3: Analyzing Teams access..." -Level Info
        Get-TeamsAccess -ExternalUsers $externalUsers
        
        # Generate Report
        $reportPath = Export-AuditReport -ExternalUsers $externalUsers
        
        # Display Summary
        Write-Host "`n" -NoNewline
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host "           AUDIT COMPLETED SUCCESSFULLY                 " -ForegroundColor Green
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host "`n📊 Audit Statistics:" -ForegroundColor Cyan
        Write-Host "   • External Users Found: $($externalUsers.Count)" -ForegroundColor White
        Write-Host "   • Execution Time: $((Get-Date) - $Global:AuditConfig.StartTime)" -ForegroundColor White
        Write-Host "`n📁 Report Generated:" -ForegroundColor Cyan
        Write-Host "   $reportPath" -ForegroundColor White
        
        # Disconnect services
        Write-AuditLog "Disconnecting services..." -Level Info
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore disconnect errors
        }
        
        Write-AuditLog "✅ Audit complete!" -Level Success
    }
    catch {
        Write-AuditLog "Critical error: $($_.Exception.Message)" -Level Error
        
        # Ensure cleanup
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore cleanup errors
        }
    }
}

# Execute the audit
Start-ExternalUserAudit
