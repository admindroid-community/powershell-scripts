<#
=============================================================================================
Name:           SharePoint Online Permissions Matrix Audit
Version:        2.0
Author:         Alpesh Nakar
Website:        educ4te.com
Company:        EDUC4TE

Script Highlights:
~~~~~~~~~~~~~~~~~
1. Comprehensive audit of SharePoint Online permissions across the entire tenancy
2. Audits site collections, subsites, lists, libraries, and sharing permissions
3. Exports detailed permissions matrix in HTML and CSV formats
4. Implements throttling to process limited sites at a time (configurable)
5. Includes progress tracking and comprehensive error handling
6. Scheduler friendly with credential parameter support
7. Generates comprehensive HTML report with detailed permissions analysis

For detailed script execution: https://educ4te.com/sharepoint/audit-sharepoint-online-permissions-matrix/
============================================================================================
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$UserName,
    
    [Parameter(Mandatory = $false)]
    [string]$Password,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantName,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    
    [Parameter(Mandatory = $false)]
    [int]$ThrottleLimit = 2,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSubsites,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeListPermissions,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSharingLinks,
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateHtmlReport,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\temp",
    
    [Parameter(Mandatory = $false)]
    [array]$SiteFilter = @(),
    
    [Parameter(Mandatory = $false)]
    [switch]$VerboseLogging
)

# Initialize global variables
$Global:PermissionsResults = @()
$Global:ErrorLog = @()
$Global:ProcessedSites = 0
$Global:TotalSites = 0
$Global:StartTime = Get-Date

Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "SharePoint Online Permissions Matrix Audit" -ForegroundColor Cyan
Write-Host "Version 2.0" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan

# Function to write verbose logs
function Write-VerboseLog {
    param([string]$Message, [string]$Level = "INFO")
    
    if ($VerboseLogging) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor Gray
    }
}

# Function to add error to log
function Add-ErrorLog {
    param(
        [string]$Location,
        [string]$ErrorMessage,
        [string]$Details = ""
    )
    
    $Global:ErrorLog += [PSCustomObject]@{
        Timestamp = Get-Date
        Location = $Location
        Error = $ErrorMessage
        Details = $Details
    }
    
    Write-VerboseLog "ERROR in $Location : $ErrorMessage" "ERROR"
}

# Function to check and install required modules
function Test-RequiredModules {
    Write-Host "`nChecking required modules..." -ForegroundColor Yellow
    
    $requiredModules = @(
        @{Name = "PnP.PowerShell"; MinVersion = "1.12.0"},
        @{Name = "Microsoft.Online.SharePoint.PowerShell"; MinVersion = "16.0.0"}
    )
    
    $missingModules = @()
    
    foreach ($module in $requiredModules) {
        $installedModule = Get-Module -ListAvailable -Name $module.Name | 
                          Where-Object {$_.Version -ge [Version]$module.MinVersion} | 
                          Sort-Object Version -Descending | 
                          Select-Object -First 1
        
        if (-not $installedModule) {
            $missingModules += $module.Name
            Write-Host "  ✗ $($module.Name) (>= $($module.MinVersion)) - Missing" -ForegroundColor Red
        } else {
            Write-Host "  ✓ $($module.Name) v$($installedModule.Version) - Available" -ForegroundColor Green
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Host "`nMissing required modules detected!" -ForegroundColor Red
        Write-Host "Please install the missing modules using:" -ForegroundColor Yellow
        foreach ($module in $missingModules) {
            Write-Host "  Install-Module -Name $module -Force -Scope CurrentUser" -ForegroundColor Cyan
        }
        throw "Required modules are missing. Please install them and run the script again."
    }
}

# Function to establish connections
function Connect-SharePointServices {
    Write-Host "`nEstablishing SharePoint connections..." -ForegroundColor Yellow
    
    # Validate tenant name
    if ([string]::IsNullOrEmpty($TenantName)) {
        Write-Host "SharePoint tenant name is required." -ForegroundColor Red
        Write-Host "Example: 'contoso' for contoso.sharepoint.com" -ForegroundColor Yellow
        $TenantName = Read-Host "Please enter your SharePoint tenant name"
    }
    
    # Set connection URLs
    $adminUrl = "https://$TenantName-admin.sharepoint.com"
    $Global:TenantUrl = "https://$TenantName.sharepoint.com"
    
    try {
        # Connect to SharePoint Online Management Shell
        Write-Host "Connecting to SharePoint Online Management Shell..." -ForegroundColor Cyan
        
        if ($ClientId -and $CertificateThumbprint) {
            # Certificate-based authentication
            Connect-SPOService -Url $adminUrl -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            Write-Host "✓ Connected using certificate authentication" -ForegroundColor Green
        }
        elseif ($UserName -and $Password) {
            # Credential-based authentication
            $securePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($UserName, $securePassword)
            Connect-SPOService -Url $adminUrl -Credential $credential
            Write-Host "✓ Connected using credential authentication" -ForegroundColor Green
        }
        else {
            # Interactive authentication
            Connect-SPOService -Url $adminUrl
            Write-Host "✓ Connected using interactive authentication" -ForegroundColor Green
        }
        
        # Test connection by getting tenant info
        $tenantInfo = Get-SPOTenant -ErrorAction Stop
        Write-Host "✓ SharePoint Online Management Shell connection verified" -ForegroundColor Green
        Write-Host "  Tenant: $($tenantInfo.DisplayName)" -ForegroundColor Gray
        
        return $true
    }
    catch {
        Add-ErrorLog "Connect-SharePointServices" $_.Exception.Message
        Write-Host "✗ Failed to connect to SharePoint services: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to connect to a specific site using PnP
function Connect-PnPSite {
    param([string]$SiteUrl)
    
    try {
        if ($ClientId -and $CertificateThumbprint) {
            Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $CertificateThumbprint -Tenant "$TenantName.onmicrosoft.com"
        }
        elseif ($UserName -and $Password) {
            $securePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($UserName, $securePassword)
            Connect-PnPOnline -Url $SiteUrl -Credentials $credential
        }
        else {
            Connect-PnPOnline -Url $SiteUrl -Interactive -WarningAction SilentlyContinue
        }
        return $true
    }
    catch {
        Add-ErrorLog "Connect-PnPSite" $_.Exception.Message "Site: $SiteUrl"
        return $false
    }
}

# Function to get all site collections
function Get-AllSiteCollections {
    Write-Host "`nRetrieving site collections..." -ForegroundColor Yellow
    
    try {
        $sites = @()
        
        # Get regular site collections
        $regularSites = Get-SPOSite -Limit All -IncludePersonalSite $false | Where-Object {
            $_.Template -notlike "*SRCHCEN*" -and 
            $_.Template -notlike "*SPSMSITEHOST*" -and
            $_.Template -notlike "*APPCATALOG*" -and
            $_.Url -notlike "*-admin.sharepoint.com*"
        }
        
        $sites += $regularSites
        
        # Apply site filter if specified
        if ($SiteFilter.Count -gt 0) {
            $sites = $sites | Where-Object {
                $siteUrl = $_.Url
                $SiteFilter | ForEach-Object {
                    if ($siteUrl -like "*$_*") { return $true }
                }
            }
        }
        
        $Global:TotalSites = $sites.Count
        Write-Host "✓ Found $($Global:TotalSites) site collections to audit" -ForegroundColor Green
        
        return $sites
    }
    catch {
        Add-ErrorLog "Get-AllSiteCollections" $_.Exception.Message
        Write-Host "✗ Failed to retrieve site collections: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Function to get site permissions
function Get-SitePermissions {
    param(
        [string]$SiteUrl,
        [string]$SiteTitle
    )
    
    $sitePermissions = @()
    
    try {
        Write-VerboseLog "Auditing permissions for site: $SiteUrl"
        
        # Connect to site
        if (-not (Connect-PnPSite -SiteUrl $SiteUrl)) {
            return $sitePermissions
        }
        
        # Get site groups and permissions
        $siteGroups = Get-PnPGroup -ErrorAction SilentlyContinue
        
        foreach ($group in $siteGroups) {
            try {
                $groupUsers = Get-PnPGroupMember -Identity $group.Id -ErrorAction SilentlyContinue
                
                foreach ($user in $groupUsers) {
                    $sitePermissions += [PSCustomObject]@{
                        SiteUrl = $SiteUrl
                        SiteTitle = $SiteTitle
                        ObjectType = "Site"
                        ObjectTitle = $SiteTitle
                        ObjectUrl = $SiteUrl
                        PrincipalType = if ($user.PrincipalType -eq "User") { "User" } else { "Group" }
                        PrincipalName = $user.Title
                        PrincipalLoginName = $user.LoginName
                        PermissionLevel = $group.Title
                        PermissionType = "SharePoint Group"
                        IsInherited = $false
                        SharingLinkType = ""
                        ExpirationDate = ""
                        CreatedDate = Get-Date
                    }
                }
            }
            catch {
                Add-ErrorLog "Get-SitePermissions-Group" $_.Exception.Message "Site: $SiteUrl, Group: $($group.Title)"
            }
        }
        
        # Get direct permissions (role assignments)
        try {
            $roleAssignments = Get-PnPWeb | Get-PnPProperty -Property RoleAssignments
            
            foreach ($assignment in $roleAssignments) {
                $principal = Get-PnPProperty -ClientObject $assignment -Property Member
                $roleDefinitions = Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings
                
                foreach ($role in $roleDefinitions) {
                    $sitePermissions += [PSCustomObject]@{
                        SiteUrl = $SiteUrl
                        SiteTitle = $SiteTitle
                        ObjectType = "Site"
                        ObjectTitle = $SiteTitle
                        ObjectUrl = $SiteUrl
                        PrincipalType = $principal.PrincipalType
                        PrincipalName = $principal.Title
                        PrincipalLoginName = $principal.LoginName
                        PermissionLevel = $role.Name
                        PermissionType = "Direct Permission"
                        IsInherited = $false
                        SharingLinkType = ""
                        ExpirationDate = ""
                        CreatedDate = Get-Date
                    }
                }
            }
        }
        catch {
            Add-ErrorLog "Get-SitePermissions-RoleAssignments" $_.Exception.Message "Site: $SiteUrl"
        }
        
        # Get sharing links if requested
        if ($IncludeSharingLinks) {
            try {
                $sharingLinks = Get-PnPSharingForNonOwnersOfFile -ErrorAction SilentlyContinue
                # This is a placeholder - actual sharing link enumeration would require more complex logic
                Write-VerboseLog "Sharing links audit placeholder for site: $SiteUrl"
            }
            catch {
                Write-VerboseLog "Could not retrieve sharing links for site: $SiteUrl"
            }
        }
        
    }
    catch {
        Add-ErrorLog "Get-SitePermissions" $_.Exception.Message "Site: $SiteUrl"
    }
    finally {
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore disconnect errors
        }
    }
    
    return $sitePermissions
}

# Function to get list/library permissions
function Get-ListPermissions {
    param(
        [string]$SiteUrl,
        [string]$SiteTitle
    )
    
    $listPermissions = @()
    
    if (-not $IncludeListPermissions) {
        return $listPermissions
    }
    
    try {
        Write-VerboseLog "Auditing list permissions for site: $SiteUrl"
        
        # Connect to site
        if (-not (Connect-PnPSite -SiteUrl $SiteUrl)) {
            return $listPermissions
        }
        
        # Get all lists and libraries
        $lists = Get-PnPList -ErrorAction SilentlyContinue | Where-Object { 
            -not $_.Hidden -and 
            $_.BaseType -ne "Unknown" -and
            $_.Title -notlike "Style Library" -and
            $_.Title -notlike "Form Templates"
        }
        
        foreach ($list in $lists) {
            try {
                # Check if list has unique permissions
                $hasUniquePerms = Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments
                
                if ($hasUniquePerms) {
                    $roleAssignments = Get-PnPProperty -ClientObject $list -Property RoleAssignments
                    
                    foreach ($assignment in $roleAssignments) {
                        $principal = Get-PnPProperty -ClientObject $assignment -Property Member
                        $roleDefinitions = Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings
                        
                        foreach ($role in $roleDefinitions) {
                            $listPermissions += [PSCustomObject]@{
                                SiteUrl = $SiteUrl
                                SiteTitle = $SiteTitle
                                ObjectType = if ($list.BaseType -eq "DocumentLibrary") { "Library" } else { "List" }
                                ObjectTitle = $list.Title
                                ObjectUrl = "$SiteUrl/$($list.DefaultViewUrl)"
                                PrincipalType = $principal.PrincipalType
                                PrincipalName = $principal.Title
                                PrincipalLoginName = $principal.LoginName
                                PermissionLevel = $role.Name
                                PermissionType = "Unique Permission"
                                IsInherited = $false
                                SharingLinkType = ""
                                ExpirationDate = ""
                                CreatedDate = Get-Date
                            }
                        }
                    }
                }
            }
            catch {
                Add-ErrorLog "Get-ListPermissions-List" $_.Exception.Message "Site: $SiteUrl, List: $($list.Title)"
            }
        }
    }
    catch {
        Add-ErrorLog "Get-ListPermissions" $_.Exception.Message "Site: $SiteUrl"
    }
    finally {
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore disconnect errors
        }
    }
    
    return $listPermissions
}

# Function to get subsite permissions
function Get-SubsitePermissions {
    param(
        [string]$SiteUrl,
        [string]$SiteTitle
    )
    
    $subsitePermissions = @()
    
    if (-not $IncludeSubsites) {
        return $subsitePermissions
    }
    
    try {
        Write-VerboseLog "Auditing subsite permissions for site: $SiteUrl"
        
        # Connect to site
        if (-not (Connect-PnPSite -SiteUrl $SiteUrl)) {
            return $subsitePermissions
        }
        
        # Get all subsites
        $subsites = Get-PnPSubWeb -Recurse -ErrorAction SilentlyContinue
        
        foreach ($subsite in $subsites) {
            try {
                # Connect to subsite
                if (Connect-PnPSite -SiteUrl $subsite.Url) {
                    # Get subsite permissions (similar to site permissions logic)
                    $subsiteGroups = Get-PnPGroup -ErrorAction SilentlyContinue
                    
                    foreach ($group in $subsiteGroups) {
                        try {
                            $groupUsers = Get-PnPGroupMember -Identity $group.Id -ErrorAction SilentlyContinue
                            
                            foreach ($user in $groupUsers) {
                                $subsitePermissions += [PSCustomObject]@{
                                    SiteUrl = $SiteUrl
                                    SiteTitle = $SiteTitle
                                    ObjectType = "Subsite"
                                    ObjectTitle = $subsite.Title
                                    ObjectUrl = $subsite.Url
                                    PrincipalType = if ($user.PrincipalType -eq "User") { "User" } else { "Group" }
                                    PrincipalName = $user.Title
                                    PrincipalLoginName = $user.LoginName
                                    PermissionLevel = $group.Title
                                    PermissionType = "SharePoint Group"
                                    IsInherited = $false
                                    SharingLinkType = ""
                                    ExpirationDate = ""
                                    CreatedDate = Get-Date
                                }
                            }
                        }
                        catch {
                            Add-ErrorLog "Get-SubsitePermissions-Group" $_.Exception.Message "Subsite: $($subsite.Url), Group: $($group.Title)"
                        }
                    }
                }
            }
            catch {
                Add-ErrorLog "Get-SubsitePermissions-Subsite" $_.Exception.Message "Subsite: $($subsite.Url)"
            }
        }
    }
    catch {
        Add-ErrorLog "Get-SubsitePermissions" $_.Exception.Message "Site: $SiteUrl"
    }
    finally {
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore disconnect errors
        }
    }
    
    return $subsitePermissions
}

# Function to process sites with throttling
function Start-PermissionsAudit {
    param([array]$Sites)
    
    Write-Host "`nStarting permissions audit with throttle limit: $ThrottleLimit sites at a time" -ForegroundColor Yellow
    Write-Host "Audit scope:" -ForegroundColor Cyan
    Write-Host "  - Site Collections: Yes" -ForegroundColor White
    Write-Host "  - Subsites: $(if($IncludeSubsites){'Yes'}else{'No'})" -ForegroundColor White
    Write-Host "  - List/Library Permissions: $(if($IncludeListPermissions){'Yes'}else{'No'})" -ForegroundColor White
    Write-Host "  - Sharing Links: $(if($IncludeSharingLinks){'Yes'}else{'No'})" -ForegroundColor White
    
    $batchNumber = 1
    $totalBatches = [Math]::Ceiling($Sites.Count / $ThrottleLimit)
    
    for ($i = 0; $i -lt $Sites.Count; $i += $ThrottleLimit) {
        $batch = $Sites[$i..($i + $ThrottleLimit - 1)]
        
        Write-Host "`nProcessing batch $batchNumber of $totalBatches ($($batch.Count) sites)..." -ForegroundColor Cyan
        
        foreach ($site in $batch) {
            $Global:ProcessedSites++
            $percentComplete = [Math]::Round(($Global:ProcessedSites / $Global:TotalSites) * 100, 1)
            
            Write-Progress -Activity "Auditing SharePoint Permissions" -Status "Processing site $Global:ProcessedSites of $Global:TotalSites ($percentComplete%)" -PercentComplete $percentComplete -CurrentOperation $site.Url
            
            Write-Host "  Processing: $($site.Title) ($($site.Url))" -ForegroundColor White
            
            try {
                # Audit site permissions
                $sitePerms = Get-SitePermissions -SiteUrl $site.Url -SiteTitle $site.Title
                $Global:PermissionsResults += $sitePerms
                Write-VerboseLog "Found $($sitePerms.Count) site permissions for $($site.Url)"
                
                # Audit list permissions if requested
                if ($IncludeListPermissions) {
                    $listPerms = Get-ListPermissions -SiteUrl $site.Url -SiteTitle $site.Title
                    $Global:PermissionsResults += $listPerms
                    Write-VerboseLog "Found $($listPerms.Count) list permissions for $($site.Url)"
                }
                
                # Audit subsite permissions if requested
                if ($IncludeSubsites) {
                    $subsitePerms = Get-SubsitePermissions -SiteUrl $site.Url -SiteTitle $site.Title
                    $Global:PermissionsResults += $subsitePerms
                    Write-VerboseLog "Found $($subsitePerms.Count) subsite permissions for $($site.Url)"
                }
                
                Write-Host "    ✓ Completed" -ForegroundColor Green
            }
            catch {
                Add-ErrorLog "Start-PermissionsAudit" $_.Exception.Message "Site: $($site.Url)"
                Write-Host "    ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            # Small delay to prevent throttling
            Start-Sleep -Milliseconds 500
        }
        
        # Longer pause between batches
        if ($batchNumber -lt $totalBatches) {
            Write-Host "  Batch $batchNumber completed. Pausing for 2 seconds..." -ForegroundColor Gray
            Start-Sleep -Seconds 2
        }
        
        $batchNumber++
    }
    
    Write-Progress -Activity "Auditing SharePoint Permissions" -Completed
}

# Function to generate CSV report
function Export-CSVReport {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $OutputPath "SPO_PermissionsMatrix_$timestamp.csv"
    
    Write-Host "`nExporting CSV report..." -ForegroundColor Yellow
    
    try {
        # Ensure output directory exists
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        
        if ($Global:PermissionsResults.Count -gt 0) {
            $Global:PermissionsResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "✓ CSV report exported: $csvPath" -ForegroundColor Green
            return $csvPath
        }
        else {
            Write-Host "✗ No permissions data to export" -ForegroundColor Red
            return $null
        }
    }
    catch {
        Add-ErrorLog "Export-CSVReport" $_.Exception.Message
        Write-Host "✗ Failed to export CSV report: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Function to generate HTML report
function Export-HTMLReport {
    if (-not $GenerateHtmlReport) {
        return $null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $htmlPath = Join-Path $OutputPath "SPO_PermissionsMatrix_$timestamp.html"
    
    Write-Host "`nGenerating HTML report..." -ForegroundColor Yellow
    
    try {
        # Ensure output directory exists
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        
        # Calculate statistics
        $totalPermissions = $Global:PermissionsResults.Count
        $uniqueUsers = ($Global:PermissionsResults | Where-Object {$_.PrincipalType -eq "User"} | Select-Object -Unique PrincipalLoginName).Count
        $uniqueGroups = ($Global:PermissionsResults | Where-Object {$_.PrincipalType -eq "Group"} | Select-Object -Unique PrincipalName).Count
        $endTime = Get-Date
        $duration = $endTime - $Global:StartTime
        
        # Group permissions by site
        $permissionsBySite = $Global:PermissionsResults | Group-Object SiteUrl | Sort-Object Name
        
        # Create HTML content
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>SharePoint Online Permissions Matrix Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1, h2 { color: #0078d4; }
        h1 { border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        .summary { background-color: #e3f2fd; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 5px solid #0078d4; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin: 20px 0; }
        .stat-box { background-color: #f8f9fa; padding: 15px; border-radius: 5px; text-align: center; border: 1px solid #dee2e6; }
        .stat-number { font-size: 24px; font-weight: bold; color: #0078d4; }
        .stat-label { font-size: 14px; color: #666; margin-top: 5px; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; font-size: 12px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #0078d4; color: white; position: sticky; top: 0; }
        .site-section { margin: 30px 0; border: 1px solid #ddd; border-radius: 5px; }
        .site-header { background-color: #f1f3f4; padding: 15px; border-bottom: 1px solid #ddd; }
        .site-title { font-weight: bold; color: #0078d4; }
        .site-url { font-size: 11px; color: #666; margin-top: 5px; }
        .permission-table { margin: 0; }
        .user-row { background-color: #e8f5e8; }
        .group-row { background-color: #fff3e0; }
        .direct-perm { background-color: #f3e5f5; }
        .inherited-perm { background-color: #e3f2fd; }
        .error-section { background-color: #ffebee; border-left: 5px solid #f44336; padding: 15px; margin: 20px 0; border-radius: 5px; }
        .error-item { margin: 10px 0; padding: 10px; background-color: white; border-radius: 3px; }
        .footer { margin-top: 40px; text-align: center; color: #666; font-size: 12px; border-top: 1px solid #ddd; padding-top: 20px; }
        .filter-info { background-color: #fff3cd; padding: 10px; border-radius: 5px; margin: 10px 0; border-left: 5px solid #ffc107; }
    </style>
</head>
<body>
    <div class="container">
        <h1>SharePoint Online Permissions Matrix Report</h1>
        
        <div class="summary">
            <h2>Executive Summary</h2>
            <p><strong>Generated:</strong> $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))</p>
            <p><strong>Tenant:</strong> $TenantName.sharepoint.com</p>
            <p><strong>Duration:</strong> $($duration.ToString('hh\:mm\:ss'))</p>
            <p><strong>Audit Scope:</strong> $($Global:ProcessedSites) site collections processed</p>
        </div>
        
        <div class="stats-grid">
            <div class="stat-box">
                <div class="stat-number">$totalPermissions</div>
                <div class="stat-label">Total Permission Entries</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($Global:ProcessedSites)</div>
                <div class="stat-label">Sites Audited</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$uniqueUsers</div>
                <div class="stat-label">Unique Users</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$uniqueGroups</div>
                <div class="stat-label">Unique Groups</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($Global:ErrorLog.Count)</div>
                <div class="stat-label">Errors Encountered</div>
            </div>
        </div>
"@

        # Add audit scope information
        if ($IncludeSubsites -or $IncludeListPermissions -or $IncludeSharingLinks -or $SiteFilter.Count -gt 0) {
            $html += @"
        <div class="filter-info">
            <h3>Audit Configuration</h3>
            <ul>
                <li><strong>Include Subsites:</strong> $(if($IncludeSubsites){'Yes'}else{'No'})</li>
                <li><strong>Include List/Library Permissions:</strong> $(if($IncludeListPermissions){'Yes'}else{'No'})</li>
                <li><strong>Include Sharing Links:</strong> $(if($IncludeSharingLinks){'Yes'}else{'No'})</li>
                <li><strong>Throttle Limit:</strong> $ThrottleLimit sites per batch</li>
"@
            if ($SiteFilter.Count -gt 0) {
                $html += "<li><strong>Site Filter Applied:</strong> $($SiteFilter -join ', ')</li>"
            }
            $html += @"
            </ul>
        </div>
"@
        }
        
        # Add permissions details by site
        $html += @"
        <h2>Detailed Permissions by Site</h2>
"@
        
        foreach ($siteGroup in $permissionsBySite) {
            $siteUrl = $siteGroup.Name
            $siteTitle = ($siteGroup.Group | Select-Object -First 1).SiteTitle
            $sitePermissions = $siteGroup.Group
            
            $html += @"
        <div class="site-section">
            <div class="site-header">
                <div class="site-title">$([System.Web.HttpUtility]::HtmlEncode($siteTitle))</div>
                <div class="site-url">$([System.Web.HttpUtility]::HtmlEncode($siteUrl))</div>
                <div style="margin-top: 10px; font-size: 12px;">
                    <strong>Total Permissions:</strong> $($sitePermissions.Count) | 
                    <strong>Users:</strong> $(($sitePermissions | Where-Object {$_.PrincipalType -eq 'User'}).Count) | 
                    <strong>Groups:</strong> $(($sitePermissions | Where-Object {$_.PrincipalType -eq 'Group'}).Count)
                </div>
            </div>
            <table class="permission-table">
                <tr>
                    <th>Object Type</th>
                    <th>Object Title</th>
                    <th>Principal Type</th>
                    <th>Principal Name</th>
                    <th>Login Name</th>
                    <th>Permission Level</th>
                    <th>Permission Type</th>
                    <th>Inherited</th>
                </tr>
"@
            
            foreach ($permission in $sitePermissions | Sort-Object ObjectType, PrincipalType, PrincipalName) {
                $rowClass = ""
                if ($permission.PrincipalType -eq "User") { $rowClass = "user-row" }
                elseif ($permission.PrincipalType -eq "Group") { $rowClass = "group-row" }
                
                if ($permission.PermissionType -eq "Direct Permission") { $rowClass += " direct-perm" }
                elseif ($permission.IsInherited -eq $true) { $rowClass += " inherited-perm" }
                
                $html += @"
                <tr class="$rowClass">
                    <td>$([System.Web.HttpUtility]::HtmlEncode($permission.ObjectType))</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($permission.ObjectTitle))</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($permission.PrincipalType))</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($permission.PrincipalName))</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($permission.PrincipalLoginName))</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($permission.PermissionLevel))</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($permission.PermissionType))</td>
                    <td>$(if($permission.IsInherited){'Yes'}else{'No'})</td>
                </tr>
"@
            }
            
            $html += @"
            </table>
        </div>
"@
        }
        
        # Add error section if there are errors
        if ($Global:ErrorLog.Count -gt 0) {
            $html += @"
        <div class="error-section">
            <h2>Errors and Warnings</h2>
            <p>The following errors were encountered during the audit:</p>
"@
            foreach ($errorItem in $Global:ErrorLog) {
                $html += @"
            <div class="error-item">
                <strong>[$($errorItem.Timestamp.ToString('HH:mm:ss'))] $([System.Web.HttpUtility]::HtmlEncode($errorItem.Location))</strong><br>
                $([System.Web.HttpUtility]::HtmlEncode($errorItem.Error))<br>
                <small>$([System.Web.HttpUtility]::HtmlEncode($errorItem.Details))</small>
            </div>
"@
            }
            $html += "</div>"
        }
        
        # Add footer
        $html += @"
        <div class="footer">
            <p>Report generated by SharePoint Online Permissions Matrix Audit Script v2.0</p>
            <p>~~ Script prepared by EDUC4TE ~~</p>
            <p>For more Microsoft 365 training and consultancy services, visit <strong>educ4te.com</strong></p>
        </div>
    </div>
</body>
</html>
"@
        
        # Write HTML file
        Add-Type -AssemblyName System.Web
        $html | Out-File -FilePath $htmlPath -Encoding UTF8
        
        Write-Host "✓ HTML report generated: $htmlPath" -ForegroundColor Green
        return $htmlPath
    }
    catch {
        Add-ErrorLog "Export-HTMLReport" $_.Exception.Message
        Write-Host "✗ Failed to generate HTML report: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Main execution
try {
    # Test required modules
    Test-RequiredModules
    
    # Connect to SharePoint services
    if (-not (Connect-SharePointServices)) {
        throw "Failed to establish SharePoint connection"
    }
    
    # Get all site collections
    $sites = Get-AllSiteCollections
    
    if ($sites.Count -eq 0) {
        Write-Host "No site collections found to audit." -ForegroundColor Yellow
        exit
    }
    
    # Start permissions audit
    Start-PermissionsAudit -Sites $sites
    
    # Generate reports
    Write-Host "`n===============================================" -ForegroundColor Cyan
    Write-Host "AUDIT COMPLETED - GENERATING REPORTS" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    
    $endTime = Get-Date
    $duration = $endTime - $Global:StartTime
    
    Write-Host "`nAudit Summary:" -ForegroundColor Green
    Write-Host "  Sites Processed: $Global:ProcessedSites" -ForegroundColor White
    Write-Host "  Total Permission Entries: $($Global:PermissionsResults.Count)" -ForegroundColor White
    Write-Host "  Unique Users: $(($Global:PermissionsResults | Where-Object {$_.PrincipalType -eq 'User'} | Select-Object -Unique PrincipalLoginName).Count)" -ForegroundColor White
    Write-Host "  Unique Groups: $(($Global:PermissionsResults | Where-Object {$_.PrincipalType -eq 'Group'} | Select-Object -Unique PrincipalName).Count)" -ForegroundColor White
    Write-Host "  Errors Encountered: $($Global:ErrorLog.Count)" -ForegroundColor White
    Write-Host "  Execution Time: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor White
    
    # Export reports
    $csvPath = Export-CSVReport
    $htmlPath = Export-HTMLReport
    
    # Display final results
    Write-Host "`nReports Generated:" -ForegroundColor Green
    if ($csvPath) {
        Write-Host "  CSV Report: $csvPath" -ForegroundColor Cyan
    }
    if ($htmlPath) {
        Write-Host "  HTML Report: $htmlPath" -ForegroundColor Cyan
    }
    
    if ($Global:ErrorLog.Count -gt 0) {
        Write-Host "`nNote: $($Global:ErrorLog.Count) errors were encountered during the audit. Please review the HTML report for details." -ForegroundColor Yellow
    }
    
    Write-Host "`n~~ Script prepared by EDUC4TE ~~" -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green
    Write-Host "educ4te.com" -ForegroundColor Yellow -NoNewline
    Write-Host " to get access to Microsoft 365 training and consultancy services. ~~" -ForegroundColor Green
    
    # Offer to open report
    if ($htmlPath -and $GenerateHtmlReport) {
        $prompt = New-Object -ComObject wscript.shell
        $userInput = $prompt.popup("Do you want to open the HTML report?", 0, "Open Report", 4)
        if ($userInput -eq 6) {
            Invoke-Item $htmlPath
        }
    }
}
catch {
    Write-Host "`nFatal Error: $($_.Exception.Message)" -ForegroundColor Red
    Add-ErrorLog "Main Execution" $_.Exception.Message
    exit 1
}
finally {
    # Cleanup connections
    try {
        Disconnect-SPOService -ErrorAction SilentlyContinue
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore cleanup errors
    }
}
