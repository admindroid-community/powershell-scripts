<#
=============================================================================================
Name:           Audit Microsoft Security Group Access to SharePoint Sites (Enhanced v2.0)
Description:    Enhanced script that audits Microsoft Security Group permissions across all SharePoint sites with detailed permission matrix
Version:        2.0
Website:        educ4te.com

Script Highlights:
~~~~~~~~~~~~~~~~~
1. Uses modern authentication to connect to SharePoint Online
2. Supports certificate-based authentication (CBA) for automation scenarios  
3. Supports MFA-enabled account authentication
4. Automatically installs required PnP.PowerShell module upon confirmation
5. Audits security group permissions across all SharePoint sites, lists, and libraries
6. Exports detailed permission matrix report matching AdminDroid format
7. Identifies direct and inherited permissions with inheritance details
8. Shows granular permission levels (Full Control, Design, Edit, Contribute, Read, etc.)
9. Includes site collections, subsites, lists, and document libraries analysis
10. Provides detailed permission breakdown with "Given through" information
11. Scheduler-friendly design for automated auditing
12. Comprehensive error handling and progress reporting
13. Enhanced output format matching enterprise audit requirements

For detailed script execution: https://educ4te.com/
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [string]$SecurityGroupName,
    [string]$TenantName,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password,
    [switch]$IncludeSubsites,
    [switch]$IncludeSystemSites,
    [switch]$IncludeListsAndLibraries
)

# Global variables for tracking
$Global:ProcessedSites = 0
$Global:TotalSites = 0
$Global:FoundPermissions = 0
$Global:ErrorLog = @()
$Global:WarningLog = @()
$Global:ClientId = ""
$Global:TenantName = ""
$Global:UserName = ""
$Global:Password = ""
$Global:CertificateThumbprint = ""
$Global:PermissionResults = @()

Function Install-RequiredModules {
    Write-Host "Checking for required PowerShell modules..." -ForegroundColor Cyan
    
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
}

Function Connect-SharePointOnline {
    Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
    
    # Set global variables for reuse
    $Global:ClientId = $ClientId
    $Global:TenantName = $TenantName
    $Global:UserName = $UserName
    $Global:Password = $Password
    $Global:CertificateThumbprint = $CertificateThumbprint
    
    # Determine admin URL
    if ($Global:TenantName -ne "") {
        $AdminUrl = "https://$Global:TenantName-admin.sharepoint.com"
    } else {
        $Global:TenantName = Read-Host "Enter your tenant name (e.g., contoso for contoso.sharepoint.com)"
        $AdminUrl = "https://$Global:TenantName-admin.sharepoint.com"
    }
    
    try {
        # Connect to SharePoint Online using PnP
        if ($Global:ClientId -ne "" -and $Global:CertificateThumbprint -ne "" -and $Global:TenantName -ne "") {
            # Certificate-based authentication
            Write-Host "Connecting using certificate-based authentication..." -ForegroundColor Green
            Connect-PnPOnline -Url $AdminUrl -ClientId $Global:ClientId -Thumbprint $Global:CertificateThumbprint -Tenant "$Global:TenantName.onmicrosoft.com"
            
        } elseif ($Global:UserName -ne "" -and $Global:Password -ne "") {
            # Basic authentication
            Write-Host "Connecting using basic authentication..." -ForegroundColor Green
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Global:Password -Force
            $Credential = New-Object System.Management.Automation.PSCredential $Global:UserName, $SecuredPassword
            Connect-PnPOnline -Url $AdminUrl -Credential $Credential
            
        } else {
            # Interactive authentication (default)
            Write-Host "Connecting using interactive authentication..." -ForegroundColor Green
            if ($Global:ClientId -eq "") {
                $Global:ClientId = Read-Host "ClientId is required to connect PnP PowerShell. Enter ClientId"
            }
            Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $Global:ClientId
        }
        
        Write-Host "✅ Successfully connected to SharePoint Online" -ForegroundColor Green
        
        # Add tenant root entry
        $Global:PermissionResults += [PSCustomObject]@{
            Type = "Microsoft 365 tenant"
            Name = "$Global:TenantName.sharepoint.com"
            URL = $AdminUrl
            ItemPath = ""
            Inheritance = ""
            Details = ""
            UserGroup = ""
            PrincipalType = ""
            AccountName = ""
            GivenThrough = ""
            FullControl = ""
            Design = ""
            Contribute = ""
            Edit = ""
            Read = ""
            RestrictedView = ""
            WebOnlyLimitedAccess = ""
        }
        
        # Add Sites entry
        $Global:PermissionResults += [PSCustomObject]@{
            Type = "Microsoft 365 tenant"
            Name = "Sites"
            URL = $AdminUrl
            ItemPath = ""
            Inheritance = ""
            Details = ""
            UserGroup = ""
            PrincipalType = ""
            AccountName = ""
            GivenThrough = ""
            FullControl = ""
            Design = ""
            Contribute = ""
            Edit = ""
            Read = ""
            RestrictedView = ""
            WebOnlyLimitedAccess = ""
        }
        
    } catch {
        Write-Host "❌ Error connecting to SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
        $Global:ErrorLog += "Connection Error: $($_.Exception.Message)"
        Exit
    }
}

Function Get-SharePointSites {
    Write-Host "Retrieving SharePoint sites..." -ForegroundColor Cyan
    
    try {
        # Get all site collections
        if ($IncludeSystemSites.IsPresent) {
            $Sites = Get-PnPTenantSite -IncludeOneDriveSites:$false
        } else {
            # Exclude system sites like search center, admin sites, etc.
            $Sites = Get-PnPTenantSite -IncludeOneDriveSites:$false | Where-Object {
                $_.Template -notmatch "SRCH|CENTRALADMIN|TENANTADMIN|STS#-1" -and
                $_.Url -notmatch "search|admin|portals|sites/appcatalog" -and
                $_.Status -eq "Active"
            }
        }
        
        $Global:TotalSites = $Sites.Count
        Write-Host "✅ Found $($Global:TotalSites) SharePoint sites to analyze" -ForegroundColor Green
        
        return $Sites
        
    } catch {
        Write-Host "❌ Error retrieving SharePoint sites: $($_.Exception.Message)" -ForegroundColor Red
        $Global:ErrorLog += "Site Retrieval Error: $($_.Exception.Message)"
        return @()
    }
}

Function Convert-PermissionLevelToMatrix {
    param([string]$PermissionLevels)
    
    $Matrix = @{
        FullControl = ""
        Design = ""
        Contribute = ""
        Edit = ""
        Read = ""
        RestrictedView = ""
        WebOnlyLimitedAccess = ""
    }
    
    if ($PermissionLevels -match "Full Control") { $Matrix.FullControl = "X" }
    if ($PermissionLevels -match "Design") { $Matrix.Design = "X" }
    if ($PermissionLevels -match "Contribute") { $Matrix.Contribute = "X" }
    if ($PermissionLevels -match "Edit") { $Matrix.Edit = "X" }
    if ($PermissionLevels -match "Read") { $Matrix.Read = "X" }
    if ($PermissionLevels -match "Restricted View") { $Matrix.RestrictedView = "X" }
    if ($PermissionLevels -match "Limited Access") { $Matrix.WebOnlyLimitedAccess = "X" }
    
    return $Matrix
}

Function Test-SecurityGroupPermissions {
    param(
        [string]$SiteUrl,
        [string]$SiteTitle,
        [string]$GroupDisplayName
    )
    
    try {
        # Connect to the specific site using same auth method as admin connection
        if ($Global:ClientId -ne "" -and $Global:CertificateThumbprint -ne "" -and $Global:TenantName -ne "") {
            Connect-PnPOnline -Url $SiteUrl -ClientId $Global:ClientId -Thumbprint $Global:CertificateThumbprint -Tenant "$Global:TenantName.onmicrosoft.com" -ErrorAction Stop
        } elseif ($Global:UserName -ne "" -and $Global:Password -ne "") {
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Global:Password -Force
            $Credential = New-Object System.Management.Automation.PSCredential $Global:UserName, $SecuredPassword
            Connect-PnPOnline -Url $SiteUrl -Credential $Credential -ErrorAction Stop
        } else {
            Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $Global:ClientId -ErrorAction Stop
        }
        
        # Get site permissions
        $Web = Get-PnPWeb -Includes RoleAssignments
        
        # Check if the group has permissions on the site
        foreach ($RoleAssignment in $Web.RoleAssignments) {
            $Principal = $null
            
            try {
                $Principal = Get-PnPProperty -ClientObject $RoleAssignment -Property Member -ErrorAction SilentlyContinue
                
                if ($null -ne $Principal) {
                    # Check if this is our target group (or if no specific group is specified, get all security groups)
                    $IsTargetGroup = $false
                    
                    if ($Principal.PrincipalType -eq "SecurityGroup" -or $Principal.PrincipalType -eq "SharePointGroup") {
                        if ([string]::IsNullOrEmpty($GroupDisplayName)) {
                            # If no specific group specified, include all security groups
                            $IsTargetGroup = $true
                        } else {
                            # Case-insensitive matching for specific group names
                            if ($Principal.Title -ieq $GroupDisplayName -or 
                                $Principal.LoginName -ilike "*$GroupDisplayName*" -or
                                $Principal.Title -ilike "*$GroupDisplayName*" -or
                                $Principal.LoginName -ieq $GroupDisplayName) {
                                $IsTargetGroup = $true
                            }
                        }
                    }
                    
                    if ($IsTargetGroup) {
                        $RoleDefinitions = Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings
                        $PermissionLevels = ($RoleDefinitions | ForEach-Object { $_.Name }) -join ", "
                        
                        # Convert permissions to matrix format
                        $PermissionMatrix = Convert-PermissionLevelToMatrix -PermissionLevels $PermissionLevels
                        
                        # Determine inheritance
                        $Inheritance = if ($Web.HasUniqueRoleAssignments) { "Custom" } else { "Inherited" }
                        
                        # Find which SharePoint group this security group belongs to
                        $GivenThrough = ""
                        try {
                            $SharePointGroups = Get-PnPGroup
                            foreach ($SPGroup in $SharePointGroups) {
                                try {
                                    $GroupUsers = Get-PnPGroupUser -Identity $SPGroup.LoginName -ErrorAction SilentlyContinue
                                    if ($GroupUsers | Where-Object { $_.LoginName -eq $Principal.LoginName }) {
                                        $GivenThrough = $SPGroup.Title
                                        break
                                    }
                                } catch {
                                    # Continue if we can't read group membership
                                }
                            }
                        } catch {
                            # If we can't determine the SharePoint group, leave it empty
                        }
                        
                        if ([string]::IsNullOrEmpty($GivenThrough)) {
                            $GivenThrough = $Principal.Title
                        }
                        
                        $Global:PermissionResults += [PSCustomObject]@{
                            Type = "Site collection"
                            Name = $SiteTitle
                            URL = $SiteUrl
                            ItemPath = ""
                            Inheritance = $Inheritance
                            Details = ""
                            UserGroup = $Principal.Title
                            PrincipalType = "Security group"
                            AccountName = $Principal.LoginName
                            GivenThrough = $GivenThrough
                            FullControl = $PermissionMatrix.FullControl
                            Design = $PermissionMatrix.Design
                            Contribute = $PermissionMatrix.Contribute
                            Edit = $PermissionMatrix.Edit
                            Read = $PermissionMatrix.Read
                            RestrictedView = $PermissionMatrix.RestrictedView
                            WebOnlyLimitedAccess = $PermissionMatrix.WebOnlyLimitedAccess
                        }
                        
                        $Global:FoundPermissions++
                    }
                }
            } catch {
                # Skip permission if we can't read it (likely insufficient permissions)
                $Global:WarningLog += "Warning: Could not read permission for $($SiteUrl)"
            }
        }
        
        # If including lists and libraries, analyze them too
        if ($IncludeListsAndLibraries.IsPresent) {
            Analyze-ListsAndLibraries -SiteUrl $SiteUrl -SiteTitle $SiteTitle -GroupDisplayName $GroupDisplayName
        }
        
        # If including subsites, check them too
        if ($IncludeSubsites.IsPresent) {
            try {
                $Subwebs = Get-PnPSubWeb -Recurse
                foreach ($Subweb in $Subwebs) {
                    Test-SecurityGroupPermissions -SiteUrl $Subweb.Url -SiteTitle $Subweb.Title -GroupDisplayName $GroupDisplayName
                }
            } catch {
                $Global:WarningLog += "Warning: Could not analyze subsites for $($SiteUrl)"
            }
        }
        
    } catch {
        $Global:ErrorLog += "Error analyzing site $($SiteUrl): $($_.Exception.Message)"
    }
}

Function Analyze-ListsAndLibraries {
    param(
        [string]$SiteUrl,
        [string]$SiteTitle,
        [string]$GroupDisplayName
    )
    
    try {
        # Get all lists and libraries
        $Lists = Get-PnPList | Where-Object { 
            $_.Hidden -eq $false -and 
            $_.Title -notmatch "^(Style Library|Master Page Gallery|Theme Gallery|Web Part Gallery|User Information List|Workflow History)$"
        }
        
        foreach ($List in $Lists) {
            try {
                # Check if list has unique permissions
                $ListItem = Get-PnPList -Identity $List.Id -Includes HasUniqueRoleAssignments, RoleAssignments
                
                if ($ListItem.HasUniqueRoleAssignments) {
                    # List has unique permissions, analyze them
                    foreach ($RoleAssignment in $ListItem.RoleAssignments) {
                        try {
                            $Principal = Get-PnPProperty -ClientObject $RoleAssignment -Property Member -ErrorAction SilentlyContinue
                            
                            if ($null -ne $Principal) {
                                # Check if this is our target group
                                $IsTargetGroup = $false
                                
                                if ($Principal.PrincipalType -eq "SecurityGroup" -or $Principal.PrincipalType -eq "SharePointGroup") {
                                    if ([string]::IsNullOrEmpty($GroupDisplayName)) {
                                        $IsTargetGroup = $true
                                    } else {
                                        if ($Principal.Title -ieq $GroupDisplayName -or 
                                            $Principal.LoginName -ilike "*$GroupDisplayName*" -or
                                            $Principal.Title -ilike "*$GroupDisplayName*" -or
                                            $Principal.LoginName -ieq $GroupDisplayName) {
                                            $IsTargetGroup = $true
                                        }
                                    }
                                }
                                
                                if ($IsTargetGroup) {
                                    $RoleDefinitions = Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings
                                    $PermissionLevels = ($RoleDefinitions | ForEach-Object { $_.Name }) -join ", "
                                    
                                    # Convert permissions to matrix format
                                    $PermissionMatrix = Convert-PermissionLevelToMatrix -PermissionLevels $PermissionLevels
                                    
                                    # Determine list type
                                    $ListType = if ($List.BaseTemplate -eq 101) { "Document library" } else { "List" }
                                    
                                    # Find which SharePoint group this security group belongs to
                                    $GivenThrough = ""
                                    try {
                                        $SharePointGroups = Get-PnPGroup
                                        foreach ($SPGroup in $SharePointGroups) {
                                            try {
                                                $GroupUsers = Get-PnPGroupUser -Identity $SPGroup.LoginName -ErrorAction SilentlyContinue
                                                if ($GroupUsers | Where-Object { $_.LoginName -eq $Principal.LoginName }) {
                                                    $GivenThrough = $SPGroup.Title
                                                    break
                                                }
                                            } catch {
                                                # Continue if we can't read group membership
                                            }
                                        }
                                    } catch {
                                        # If we can't determine the SharePoint group, leave it empty
                                    }
                                    
                                    if ([string]::IsNullOrEmpty($GivenThrough)) {
                                        $GivenThrough = $Principal.Title
                                    }
                                    
                                    $Global:PermissionResults += [PSCustomObject]@{
                                        Type = $ListType
                                        Name = $List.Title
                                        URL = "$SiteUrl/$($List.DefaultViewUrl.TrimStart('/'))"
                                        ItemPath = ""
                                        Inheritance = "Custom"
                                        Details = ""
                                        UserGroup = $Principal.Title
                                        PrincipalType = "Security group"
                                        AccountName = $Principal.LoginName
                                        GivenThrough = $GivenThrough
                                        FullControl = $PermissionMatrix.FullControl
                                        Design = $PermissionMatrix.Design
                                        Contribute = $PermissionMatrix.Contribute
                                        Edit = $PermissionMatrix.Edit
                                        Read = $PermissionMatrix.Read
                                        RestrictedView = $PermissionMatrix.RestrictedView
                                        WebOnlyLimitedAccess = $PermissionMatrix.WebOnlyLimitedAccess
                                    }
                                    
                                    $Global:FoundPermissions++
                                }
                            }
                        } catch {
                            # Skip if we can't read the permission
                        }
                    }
                } else {
                    # List inherits permissions, add inherited entry if we found permissions at site level
                    $SiteHasGroupPermissions = $Global:PermissionResults | Where-Object { 
                        $_.URL -eq $SiteUrl -and 
                        $_.Type -eq "Site collection" -and 
                        ($_.UserGroup -ieq $GroupDisplayName -or [string]::IsNullOrEmpty($GroupDisplayName))
                    }
                    
                    if ($SiteHasGroupPermissions) {
                        $ListType = if ($List.BaseTemplate -eq 101) { "Document library" } else { "List" }
                        
                        $Global:PermissionResults += [PSCustomObject]@{
                            Type = $ListType
                            Name = $List.Title
                            URL = "$SiteUrl/$($List.DefaultViewUrl.TrimStart('/'))"
                            ItemPath = ""
                            Inheritance = "Inherited"
                            Details = ""
                            UserGroup = ""
                            PrincipalType = ""
                            AccountName = ""
                            GivenThrough = ""
                            FullControl = ""
                            Design = ""
                            Contribute = ""
                            Edit = ""
                            Read = ""
                            RestrictedView = ""
                            WebOnlyLimitedAccess = ""
                        }
                    }
                }
            } catch {
                $Global:WarningLog += "Warning: Could not analyze list $($List.Title) in $($SiteUrl)"
            }
        }
    } catch {
        $Global:WarningLog += "Warning: Could not retrieve lists for $($SiteUrl)"
    }
}

Function Export-Results {
    param([array]$Results)
    
    $OutputCSV = ".\SecurityGroupSharePointAccessMatrix_$((Get-Date -format "yyyy-MMM-dd-ddd_hh-mm-ss_tt").ToString()).csv"
    
    if ($Results.Count -gt 0) {
        $Results | Export-Csv -Path $OutputCSV -NoTypeInformation
        
        Write-Host "`n✅ The output file contains $($Results.Count) audit records" -ForegroundColor Green
        Write-Host "📁 The output file is available at: " -NoNewline -ForegroundColor Yellow
        Write-Host $OutputCSV -ForegroundColor Cyan
        
        # Show summary
        Write-Host "`n📊 AUDIT SUMMARY:" -ForegroundColor Cyan
        Write-Host "   • Total Sites Processed: $Global:ProcessedSites" -ForegroundColor White
        Write-Host "   • Total Audit Records: $($Results.Count)" -ForegroundColor White
        Write-Host "   • Security Group Permissions Found: $Global:FoundPermissions" -ForegroundColor White
        Write-Host "   • Warnings: $($Global:WarningLog.Count)" -ForegroundColor Yellow
        Write-Host "   • Errors: $($Global:ErrorLog.Count)" -ForegroundColor Red
        
        # Show permission distribution
        $PermissionStats = $Results | Where-Object { $_.UserGroup -ne "" } | Group-Object PrincipalType
        if ($PermissionStats) {
            Write-Host "`n📈 PERMISSION DISTRIBUTION:" -ForegroundColor Cyan
            foreach ($Stat in $PermissionStats) {
                Write-Host "   • $($Stat.Name): $($Stat.Count) permissions" -ForegroundColor White
            }
        }
        
        # Offer to open the file
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open the output file?", 0, "Open Output File", 4)
        if ($UserInput -eq 6) {
            Invoke-Item $OutputCSV
        }
        
    } else {
        Write-Host "`n⚠️  No permissions found for the specified criteria" -ForegroundColor Yellow
        Write-Host "   This could mean:" -ForegroundColor Yellow
        Write-Host "   • No security groups have SharePoint access" -ForegroundColor Yellow
        Write-Host "   • The specified group name doesn't match exactly" -ForegroundColor Yellow
        Write-Host "   • Insufficient permissions to read site permissions" -ForegroundColor Yellow
    }
    
    # Display errors and warnings if any
    if ($Global:ErrorLog.Count -gt 0) {
        Write-Host "`n❌ ERRORS ENCOUNTERED:" -ForegroundColor Red
        $Global:ErrorLog | ForEach-Object { Write-Host "   • $_" -ForegroundColor Red }
    }
    
    if ($Global:WarningLog.Count -gt 0) {
        Write-Host "`n⚠️  WARNINGS:" -ForegroundColor Yellow
        $Global:WarningLog | ForEach-Object { Write-Host "   • $_" -ForegroundColor Yellow }
    }
}

Function Disconnect-Services {
    try {
        Write-Host "`nDisconnecting from SharePoint Online..." -ForegroundColor Cyan
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "✅ Successfully disconnected from SharePoint Online" -ForegroundColor Green
    } catch {
        # Ignore disconnection errors
    }
}

# Main execution
Write-Host "🔍 Enhanced Microsoft Security Group SharePoint Access Audit Tool v2.0" -ForegroundColor Magenta
Write-Host "=" * 80 -ForegroundColor Magenta

try {
    # Install required modules
    Install-RequiredModules
    
    # Connect to SharePoint Online
    Connect-SharePointOnline
    
    # Get all SharePoint sites
    $Sites = Get-SharePointSites
    
    if ($Sites.Count -eq 0) {
        Write-Host "❌ No SharePoint sites found to analyze" -ForegroundColor Red
        Exit
    }
    
    # Analyze each site for group permissions
    Write-Host "`n🔍 Analyzing SharePoint sites for security group permissions..." -ForegroundColor Cyan
    
    if ([string]::IsNullOrEmpty($SecurityGroupName)) {
        Write-Host "   📋 Auditing ALL security groups across all sites" -ForegroundColor Yellow
    } else {
        Write-Host "   📋 Auditing security group: $SecurityGroupName" -ForegroundColor Yellow
    }
    
    foreach ($Site in $Sites) {
        $Global:ProcessedSites++
        $ProgressPercent = [math]::Round(($Global:ProcessedSites / $Global:TotalSites) * 100, 2)
        
        Write-Progress -Activity "Analyzing SharePoint Sites" `
                      -Status "Processing: $($Site.Title) ($Global:ProcessedSites of $Global:TotalSites)" `
                      -PercentComplete $ProgressPercent
        
        Write-Host "   Analyzing site: $($Site.Title) ($($Site.Url))" -ForegroundColor Gray
        
        try {
            Test-SecurityGroupPermissions -SiteUrl $Site.Url -SiteTitle $Site.Title -GroupDisplayName $SecurityGroupName
            
        } catch {
            $Global:ErrorLog += "Failed to analyze site: $($Site.Url) - $($_.Exception.Message)"
            Write-Host "     ❌ Error analyzing site" -ForegroundColor Red
        }
    }
    
    Write-Progress -Activity "Analyzing SharePoint Sites" -Completed
    
    # Export results
    Export-Results -Results $Global:PermissionResults
    
} catch {
    Write-Host "❌ Critical error during execution: $($_.Exception.Message)" -ForegroundColor Red
    $Global:ErrorLog += "Critical Error: $($_.Exception.Message)"
} finally {
    # Always disconnect services
    Disconnect-Services
    
    # Display script information
    Write-Host "`n~~ Enhanced Security Group Audit Tool v2.0 by educ4te.com ~~" -ForegroundColor Green
    Write-Host "   For more advanced M365 reporting tools, visit AdminDroid.com" -ForegroundColor Green
}
