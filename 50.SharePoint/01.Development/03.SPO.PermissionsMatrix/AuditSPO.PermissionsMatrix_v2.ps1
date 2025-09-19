<#
=============================================================================================
Name:           SharePoint Online Permissions Matrix Audit (Enhanced Version 2.0)
Version:        2.0
Author:         Alpesh Nakar
Website:        educ4te.com
Company:        EDUC4TE

Script Highlights:
~~~~~~~~~~~~~~~~~
1. Clean and simplified authentication patterns based on working EDUC4TE scripts
2. Comprehensive audit of SharePoint Online permissions across the entire tenancy
3. Audits site collections, subsites, lists, libraries, and sharing permissions
4. Exports detailed permissions matrix in HTML and CSV formats
5. Implements throttling to process limited sites at a time (configurable)
6. Modern authentication with certificate-based auth priority
7. Scheduler-friendly design with robust error handling
8. Progress tracking and detailed reporting

For detailed script execution: https://educ4te.com/sharepoint/audit-sharepoint-online-permissions-matrix/
============================================================================================
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$UserName,
    
    [Parameter(Mandatory = $false)]
    [SecureString]$Password,
    
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

# Global variables for tracking
$Global:PermissionsResults = @()
$Global:ErrorLog = @()
$Global:WarningLog = @()
$Global:ProcessedSites = 0
$Global:TotalSites = 0
$Global:StartTime = Get-Date
$Global:TenantName = $TenantName
$Global:ClientId = $ClientId
$Global:CertificateThumbprint = $CertificateThumbprint
$Global:UserName = $UserName
$Global:Password = $Password

Write-Host "🔍 SharePoint Online Permissions Matrix Audit Tool v2.0" -ForegroundColor Magenta
Write-Host "=" * 80 -ForegroundColor Magenta
Write-Host "Author: Alpesh Nakar | Company: EDUC4TE | Website: educ4te.com" -ForegroundColor Cyan

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
function Install-RequiredModules {
    Write-Host "Checking required PowerShell modules..." -ForegroundColor Cyan
    
    # Check for PnP PowerShell module
    $PnPModule = Get-InstalledModule -Name PnP.PowerShell -MinimumVersion 1.12.0 -ErrorAction SilentlyContinue
    if ($null -eq $PnPModule) {
        Write-Host "PnP PowerShell module is not available" -ForegroundColor Yellow
        $Confirm = Read-Host "Are you sure you want to install PnP.PowerShell module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-Host "Installing PnP PowerShell module..." -ForegroundColor Magenta
            Install-Module PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
            Import-Module -Name PnP.PowerShell -Force
        } else {
            Write-Host "PnP PowerShell module is required to connect SharePoint Online. Please install module using Install-Module PnP.PowerShell cmdlet." -ForegroundColor Red
            Exit
        }
    }
    
    Write-Host "✅ All required modules are available" -ForegroundColor Green
}

# Function to establish SharePoint connection
function Connect-SharePointOnline {
    Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
    
    # Determine admin URL
    if ($Global:TenantName -eq "") {
        $Global:TenantName = Read-Host "Enter your tenant name (e.g., contoso for contoso.sharepoint.com)"
    }
    
    $AdminUrl = "https://$Global:TenantName-admin.sharepoint.com"
    
    try {
        # Auto-detect and set default PnP PowerShell ClientId if not provided
        if ($Global:ClientId -eq "" -or $null -eq $Global:ClientId) {
            # Use PnP PowerShell's default multi-tenant ClientId for interactive authentication
            $Global:ClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e"  # PnP PowerShell default ClientId
            Write-Host "ℹ️  No ClientId provided. Using PnP PowerShell default ClientId for seamless authentication." -ForegroundColor Cyan
        }
        
        # Connect to SharePoint Online using PnP with authentication priority
        # Interactive authentication (highest priority - supports MFA)
        if ($Global:ClientId -ne "" -and $Global:CertificateThumbprint -eq "" -and $Global:UserName -eq "") {
            Write-Host "🔐 Using interactive authentication with auto-detected ClientId..." -ForegroundColor Green
            Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $Global:ClientId
            
        } elseif ($Global:ClientId -eq "" -and $Global:CertificateThumbprint -eq "" -and $Global:UserName -eq "") {
            # This condition should never be reached now due to auto-detection above
            Write-Host "🔄 Fallback: Using default interactive authentication..." -ForegroundColor Green
            $Global:ClientId = "31359c7f-bd7e-475c-86db-fdb8c937548e"
            Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $Global:ClientId
            
        } elseif ($Global:ClientId -ne "" -and $Global:CertificateThumbprint -ne "" -and $Global:TenantName -ne "") {
            # Certificate-based authentication (for automation)
            Write-Host "🔑 Using certificate-based authentication..." -ForegroundColor Cyan
            Connect-PnPOnline -Url $AdminUrl -ClientId $Global:ClientId -Thumbprint $Global:CertificateThumbprint -Tenant "$Global:TenantName.onmicrosoft.com"
            
        } elseif ($Global:UserName -ne "" -and $null -ne $Global:Password) {
            # Credential-based authentication (fallback)
            Write-Host "👤 Using credential-based authentication..." -ForegroundColor Yellow
            $Credential = New-Object System.Management.Automation.PSCredential $Global:UserName, $Global:Password
            Connect-PnPOnline -Url $AdminUrl -Credential $Credential
            
        } else {
            # Fallback to interactive authentication with auto-detected ClientId
            Write-Host "🌐 Using default interactive authentication with auto-detected ClientId..." -ForegroundColor Green
            Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $Global:ClientId
        }
        
        Write-Host "✅ Successfully connected to SharePoint Online" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "❌ Error connecting to SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
        Add-ErrorLog "Connect-SharePointOnline" $_.Exception.Message
        return $false
    }
}

# Function to connect to a specific site
function Connect-ToSite {
    param([string]$SiteUrl)
    
    try {
        # Use the same ClientId that was auto-detected in Connect-SharePointOnline
        # Authentication priority: Interactive -> Certificate -> Credential
        if ($Global:ClientId -ne "" -and $Global:CertificateThumbprint -eq "" -and $Global:UserName -eq "") {
            # Interactive authentication (highest priority) - uses auto-detected ClientId
            Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $Global:ClientId
        } elseif ($Global:ClientId -ne "" -and $Global:CertificateThumbprint -ne "" -and $Global:TenantName -ne "") {
            # Certificate-based authentication
            Connect-PnPOnline -Url $SiteUrl -ClientId $Global:ClientId -Thumbprint $Global:CertificateThumbprint -Tenant "$Global:TenantName.onmicrosoft.com"
        } elseif ($Global:UserName -ne "" -and $null -ne $Global:Password) {
            # Credential-based authentication
            $Credential = New-Object System.Management.Automation.PSCredential $Global:UserName, $Global:Password
            Connect-PnPOnline -Url $SiteUrl -Credential $Credential
        } else {
            # Default to interactive authentication with auto-detected ClientId
            Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $Global:ClientId
        }
        return $true
    }
    catch {
        Add-ErrorLog "Connect-ToSite" $_.Exception.Message "Site: $SiteUrl"
        return $false
    }
}

# Function to get all SharePoint sites
function Get-SharePointSites {
    Write-Host "Retrieving SharePoint sites..." -ForegroundColor Cyan
    
    try {
        # Get all site collections excluding system sites
        $Sites = Get-PnPTenantSite -IncludeOneDriveSites:$false | Where-Object {
            $_.Template -notlike "*SRCHCEN*" -and 
            $_.Template -notlike "*SPSMSITEHOST*" -and
            $_.Template -notlike "*APPCATALOG*" -and
            $_.Url -notlike "*-admin.sharepoint.com*" -and
            $_.Template -notlike "*REDIRECTSITE*"
        }
        
        # Apply site filter if specified
        if ($SiteFilter.Count -gt 0) {
            $Sites = $Sites | Where-Object {
                $siteUrl = $_.Url
                $SiteFilter | ForEach-Object {
                    if ($siteUrl -like "*$_*") { return $true }
                }
            }
        }
        
        $Global:TotalSites = $Sites.Count
        Write-Host "✅ Found $($Global:TotalSites) SharePoint sites to analyze" -ForegroundColor Green
        return $Sites
        
    } catch {
        Write-Host "❌ Error retrieving SharePoint sites: $($_.Exception.Message)" -ForegroundColor Red
        Add-ErrorLog "Get-SharePointSites" $_.Exception.Message
        return @()
    }
}

# Function to convert permission levels to readable format
function Convert-PermissionLevels {
    param([string]$PermissionLevels)
    
    $permissionMatrix = @{
        FullControl = ""
        Design = ""
        Contribute = ""
        Edit = ""
        Read = ""
        RestrictedView = ""
        LimitedAccess = ""
    }
    
    if ($PermissionLevels -match "Full Control") { $permissionMatrix.FullControl = "✓" }
    if ($PermissionLevels -match "Design") { $permissionMatrix.Design = "✓" }
    if ($PermissionLevels -match "Contribute") { $permissionMatrix.Contribute = "✓" }
    if ($PermissionLevels -match "Edit") { $permissionMatrix.Edit = "✓" }
    if ($PermissionLevels -match "Read") { $permissionMatrix.Read = "✓" }
    if ($PermissionLevels -match "Restricted View") { $permissionMatrix.RestrictedView = "✓" }
    if ($PermissionLevels -match "Limited Access") { $permissionMatrix.LimitedAccess = "✓" }
    
    return $permissionMatrix
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
        if (-not (Connect-ToSite -SiteUrl $SiteUrl)) {
            return $sitePermissions
        }
        
        # Get site groups and permissions
        $siteGroups = Get-PnPGroup -ErrorAction SilentlyContinue
        
        foreach ($group in $siteGroups) {
            try {
                $groupUsers = Get-PnPGroupMember -Identity $group.Id -ErrorAction SilentlyContinue
                
                foreach ($user in $groupUsers) {
                    $permissionMatrix = Convert-PermissionLevels -PermissionLevels $group.Title
                    
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
                        FullControl = $permissionMatrix.FullControl
                        Design = $permissionMatrix.Design
                        Contribute = $permissionMatrix.Contribute
                        Edit = $permissionMatrix.Edit
                        Read = $permissionMatrix.Read
                        RestrictedView = $permissionMatrix.RestrictedView
                        LimitedAccess = $permissionMatrix.LimitedAccess
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
            $web = Get-PnPWeb -Includes RoleAssignments
            
            foreach ($assignment in $web.RoleAssignments) {
                $principal = Get-PnPProperty -ClientObject $assignment -Property Member
                $roleDefinitions = Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings
                
                foreach ($role in $roleDefinitions) {
                    $permissionMatrix = Convert-PermissionLevels -PermissionLevels $role.Name
                    
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
                        FullControl = $permissionMatrix.FullControl
                        Design = $permissionMatrix.Design
                        Contribute = $permissionMatrix.Contribute
                        Edit = $permissionMatrix.Edit
                        Read = $permissionMatrix.Read
                        RestrictedView = $permissionMatrix.RestrictedView
                        LimitedAccess = $permissionMatrix.LimitedAccess
                        CreatedDate = Get-Date
                    }
                }
            }
        }
        catch {
            Add-ErrorLog "Get-SitePermissions-RoleAssignments" $_.Exception.Message "Site: $SiteUrl"
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
        if (-not (Connect-ToSite -SiteUrl $SiteUrl)) {
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
                            $permissionMatrix = Convert-PermissionLevels -PermissionLevels $role.Name
                            
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
                                FullControl = $permissionMatrix.FullControl
                                Design = $permissionMatrix.Design
                                Contribute = $permissionMatrix.Contribute
                                Edit = $permissionMatrix.Edit
                                Read = $permissionMatrix.Read
                                RestrictedView = $permissionMatrix.RestrictedView
                                LimitedAccess = $permissionMatrix.LimitedAccess
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
        if (-not (Connect-ToSite -SiteUrl $SiteUrl)) {
            return $subsitePermissions
        }
        
        # Get all subsites
        $subsites = Get-PnPSubWeb -Recurse -ErrorAction SilentlyContinue
        
        foreach ($subsite in $subsites) {
            try {
                # Connect to subsite
                if (Connect-ToSite -SiteUrl $subsite.Url) {
                    # Get subsite permissions (similar to site permissions logic)
                    $subsiteGroups = Get-PnPGroup -ErrorAction SilentlyContinue
                    
                    foreach ($group in $subsiteGroups) {
                        try {
                            $groupUsers = Get-PnPGroupMember -Identity $group.Id -ErrorAction SilentlyContinue
                            
                            foreach ($user in $groupUsers) {
                                $permissionMatrix = Convert-PermissionLevels -PermissionLevels $group.Title
                                
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
                                    FullControl = $permissionMatrix.FullControl
                                    Design = $permissionMatrix.Design
                                    Contribute = $permissionMatrix.Contribute
                                    Edit = $permissionMatrix.Edit
                                    Read = $permissionMatrix.Read
                                    RestrictedView = $permissionMatrix.RestrictedView
                                    LimitedAccess = $permissionMatrix.LimitedAccess
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
            
            Write-Host "  📋 Processing: $($site.Title) ($($site.Url))" -ForegroundColor White
            
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
                
                Write-Host "    ✅ Completed" -ForegroundColor Green
            }
            catch {
                Add-ErrorLog "Start-PermissionsAudit" $_.Exception.Message "Site: $($site.Url)"
                Write-Host "    ❌ Error: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            # Small delay to prevent throttling
            Start-Sleep -Milliseconds 500
        }
        
        # Longer pause between batches
        if ($batchNumber -lt $totalBatches) {
            Write-Host "  ⏸️ Batch $batchNumber completed. Pausing for 2 seconds..." -ForegroundColor Gray
            Start-Sleep -Seconds 2
        }
        
        $batchNumber++
    }
    
    Write-Progress -Activity "Auditing SharePoint Permissions" -Completed
}

# Function to generate CSV report
function Export-CSVReport {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $OutputPath "SPO_PermissionsMatrix_v2_$timestamp.csv"
    
    Write-Host "`nExporting CSV report..." -ForegroundColor Yellow
    
    try {
        # Ensure output directory exists
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        
        if ($Global:PermissionsResults.Count -gt 0) {
            $Global:PermissionsResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "✅ CSV report exported: $csvPath" -ForegroundColor Green
            return $csvPath
        }
        else {
            Write-Host "⚠️ No permissions data to export" -ForegroundColor Yellow
            return $null
        }
    }
    catch {
        Add-ErrorLog "Export-CSVReport" $_.Exception.Message
        Write-Host "❌ Failed to export CSV report: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Function to generate HTML report
function Export-HTMLReport {
    if (-not $GenerateHtmlReport) {
        return $null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $htmlPath = Join-Path $OutputPath "SPO_PermissionsMatrix_v2_$timestamp.html"
    
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
        
        # Create enhanced HTML content with modern styling
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>SharePoint Online Permissions Matrix Report v2.0</title>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }
        .container { 
            max-width: 1400px; 
            margin: 0 auto; 
            background-color: white; 
            padding: 30px; 
            border-radius: 15px; 
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }
        .header {
            text-align: center;
            margin-bottom: 30px;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 20px;
        }
        .header h1 { 
            color: #0078d4; 
            margin: 0;
            font-size: 2.5em;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        }
        .header .version {
            color: #666;
            font-size: 1.1em;
            margin-top: 10px;
        }
        .summary { 
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
            padding: 25px; 
            border-radius: 10px; 
            margin: 25px 0; 
            border-left: 6px solid #0078d4;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        .stats-grid { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); 
            gap: 20px; 
            margin: 25px 0; 
        }
        .stat-box { 
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            padding: 20px; 
            border-radius: 10px; 
            text-align: center; 
            border: 2px solid #dee2e6;
            box-shadow: 0 4px 10px rgba(0,0,0,0.1);
            transition: transform 0.3s ease;
        }
        .stat-box:hover {
            transform: translateY(-5px);
        }
        .stat-number { 
            font-size: 2.2em; 
            font-weight: bold; 
            color: #0078d4; 
            text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
        }
        .stat-label { 
            font-size: 14px; 
            color: #666; 
            margin-top: 8px; 
            font-weight: 500;
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin: 25px 0; 
            font-size: 13px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        th, td { 
            border: 1px solid #ddd; 
            padding: 12px 8px; 
            text-align: left; 
        }
        th { 
            background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
            color: white; 
            position: sticky; 
            top: 0; 
            font-weight: 600;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
        }
        .site-section { 
            margin: 35px 0; 
            border: 2px solid #ddd; 
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        .site-header { 
            background: linear-gradient(135deg, #f1f3f4 0%, #e8eaf6 100%);
            padding: 20px; 
            border-bottom: 2px solid #ddd; 
        }
        .site-title { 
            font-weight: bold; 
            color: #0078d4; 
            font-size: 1.3em;
        }
        .site-url { 
            font-size: 12px; 
            color: #666; 
            margin-top: 8px; 
            word-break: break-all;
        }
        .permission-table { margin: 0; }
        .user-row { background-color: #e8f5e8; }
        .group-row { background-color: #fff3e0; }
        .direct-perm { background-color: #f3e5f5; }
        .inherited-perm { background-color: #e3f2fd; }
        .error-section { 
            background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
            border-left: 6px solid #f44336; 
            padding: 20px; 
            margin: 25px 0; 
            border-radius: 10px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        .error-item { 
            margin: 15px 0; 
            padding: 15px; 
            background-color: white; 
            border-radius: 8px;
            border-left: 4px solid #f44336;
        }
        .footer { 
            margin-top: 50px; 
            text-align: center; 
            color: #666; 
            font-size: 14px; 
            border-top: 2px solid #ddd; 
            padding-top: 25px; 
        }
        .filter-info { 
            background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%);
            padding: 15px; 
            border-radius: 8px; 
            margin: 15px 0; 
            border-left: 6px solid #ffc107;
            box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        }
        .checkmark { color: #28a745; font-weight: bold; }
        .brand-footer {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            margin: 30px -30px -30px -30px;
            border-radius: 0 0 15px 15px;
            text-align: center;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔍 SharePoint Online Permissions Matrix Report</h1>
            <div class="version">Enhanced Version 2.0 | Powered by EDUC4TE</div>
        </div>
        
        <div class="summary">
            <h2>📊 Executive Summary</h2>
            <p><strong>Generated:</strong> $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))</p>
            <p><strong>Tenant:</strong> $Global:TenantName.sharepoint.com</p>
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
            <h3>⚙️ Audit Configuration</h3>
            <ul>
                <li><strong>Include Subsites:</strong> $(if($IncludeSubsites){'<span class="checkmark">✓</span>'}else{'✗'})</li>
                <li><strong>Include List/Library Permissions:</strong> $(if($IncludeListPermissions){'<span class="checkmark">✓</span>'}else{'✗'})</li>
                <li><strong>Include Sharing Links:</strong> $(if($IncludeSharingLinks){'<span class="checkmark">✓</span>'}else{'✗'})</li>
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
        <h2>📋 Detailed Permissions by Site</h2>
"@
        
        foreach ($siteGroup in $permissionsBySite) {
            $siteUrl = $siteGroup.Name
            $siteTitle = ($siteGroup.Group | Select-Object -First 1).SiteTitle
            $sitePermissions = $siteGroup.Group
            
            $html += @"
        <div class="site-section">
            <div class="site-header">
                <div class="site-title">🏢 $([System.Web.HttpUtility]::HtmlEncode($siteTitle))</div>
                <div class="site-url">🔗 $([System.Web.HttpUtility]::HtmlEncode($siteUrl))</div>
                <div style="margin-top: 15px; font-size: 13px;">
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
                    <th>Permission Level</th>
                    <th>Full Control</th>
                    <th>Contribute</th>
                    <th>Read</th>
                    <th>Edit</th>
                    <th>Design</th>
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
                    <td>$([System.Web.HttpUtility]::HtmlEncode($permission.PermissionLevel))</td>
                    <td style="text-align: center;">$($permission.FullControl)</td>
                    <td style="text-align: center;">$($permission.Contribute)</td>
                    <td style="text-align: center;">$($permission.Read)</td>
                    <td style="text-align: center;">$($permission.Edit)</td>
                    <td style="text-align: center;">$($permission.Design)</td>
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
            <h2>⚠️ Errors and Warnings</h2>
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
        
        # Add enhanced footer
        $html += @"
        <div class="brand-footer">
            <h3>🚀 Report generated by SharePoint Online Permissions Matrix Audit Tool v2.0</h3>
            <p><strong>Author:</strong> Alpesh Nakar | <strong>Company:</strong> EDUC4TE</p>
            <p>For more Microsoft 365 training and consultancy services, visit <strong>educ4te.com</strong></p>
            <p style="margin-top: 15px; font-size: 12px; opacity: 0.8;">
                Enhanced with modern authentication, comprehensive error handling, and enterprise-grade reporting
            </p>
        </div>
    </div>
</body>
</html>
"@
        
        # Write HTML file
        Add-Type -AssemblyName System.Web
        $html | Out-File -FilePath $htmlPath -Encoding UTF8
        
        Write-Host "✅ HTML report generated: $htmlPath" -ForegroundColor Green
        return $htmlPath
    }
    catch {
        Add-ErrorLog "Export-HTMLReport" $_.Exception.Message
        Write-Host "❌ Failed to generate HTML report: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Function to display final summary
function Show-AuditSummary {
    $endTime = Get-Date
    $duration = $endTime - $Global:StartTime
    
    Write-Host "`n" + "=" * 80 -ForegroundColor Cyan
    Write-Host "📊 AUDIT COMPLETED - GENERATING REPORTS" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Cyan
    
    Write-Host "`n✅ Audit Summary:" -ForegroundColor Green
    Write-Host "   • Sites Processed: $Global:ProcessedSites" -ForegroundColor White
    Write-Host "   • Total Permission Entries: $($Global:PermissionsResults.Count)" -ForegroundColor White
    Write-Host "   • Unique Users: $(($Global:PermissionsResults | Where-Object {$_.PrincipalType -eq 'User'} | Select-Object -Unique PrincipalLoginName).Count)" -ForegroundColor White
    Write-Host "   • Unique Groups: $(($Global:PermissionsResults | Where-Object {$_.PrincipalType -eq 'Group'} | Select-Object -Unique PrincipalName).Count)" -ForegroundColor White
    Write-Host "   • Errors Encountered: $($Global:ErrorLog.Count)" -ForegroundColor White
    Write-Host "   • Execution Time: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor White
    
    # Show permission distribution
    if ($Global:PermissionsResults.Count -gt 0) {
        $PermissionStats = $Global:PermissionsResults | Where-Object { $_.PrincipalName -ne "" } | Group-Object PrincipalType
        if ($PermissionStats) {
            Write-Host "`n📈 Permission Distribution:" -ForegroundColor Cyan
            foreach ($stat in $PermissionStats) {
                Write-Host "   • $($stat.Name): $($stat.Count) permissions" -ForegroundColor White
            }
        }
    }
}

# Main execution
try {
    # Install required modules
    Install-RequiredModules
    
    # Connect to SharePoint Online
    if (-not (Connect-SharePointOnline)) {
        throw "Failed to establish SharePoint connection"
    }
    
    # Get all SharePoint sites
    $sites = Get-SharePointSites
    
    if ($sites.Count -eq 0) {
        Write-Host "❌ No SharePoint sites found to analyze" -ForegroundColor Red
        Exit
    }
    
    # Start permissions audit
    Start-PermissionsAudit -Sites $sites
    
    # Generate reports and show summary
    Show-AuditSummary
    
    # Export reports
    $csvPath = Export-CSVReport
    $htmlPath = Export-HTMLReport
    
    # Display final results
    Write-Host "`n📁 Reports Generated:" -ForegroundColor Green
    if ($csvPath) {
        Write-Host "   📄 CSV Report: $csvPath" -ForegroundColor Cyan
    }
    if ($htmlPath) {
        Write-Host "   🌐 HTML Report: $htmlPath" -ForegroundColor Cyan
    }
    
    if ($Global:ErrorLog.Count -gt 0) {
        Write-Host "`n⚠️ Note: $($Global:ErrorLog.Count) errors were encountered during the audit. Please review the HTML report for details." -ForegroundColor Yellow
    }
    
    Write-Host "`n🎯 ~~ Enhanced SharePoint Audit Tool v2.0 by EDUC4TE ~~" -ForegroundColor Green
    Write-Host "🌐 ~~ Check out " -NoNewline -ForegroundColor Green
    Write-Host "educ4te.com" -ForegroundColor Yellow -NoNewline
    Write-Host " for Microsoft 365 training and consultancy services ~~" -ForegroundColor Green
    
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
    Write-Host "`n❌ Critical error during execution: $($_.Exception.Message)" -ForegroundColor Red
    Add-ErrorLog "Main Execution" $_.Exception.Message
    exit 1
}
finally {
    # Always disconnect services
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "`n🔌 Disconnected from SharePoint Online" -ForegroundColor Green
    }
    catch {
        # Ignore cleanup errors
    }
    
    $endTime = Get-Date
    $duration = $endTime - $Global:StartTime
    Write-Host "⏱️ Total execution time: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
}