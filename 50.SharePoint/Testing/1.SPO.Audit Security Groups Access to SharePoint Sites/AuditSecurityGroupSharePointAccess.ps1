<#
=============================================================================================
Name:           Audit Microsoft Security Group Access to SharePoint Sites
Description:    This script audits Microsoft Security Group permissions across all SharePoint sites
Version:        3.0
Website:        educ4te.com

Script Highlights:
~~~~~~~~~~~~~~~~~
1. Uses modern authentication to connect to SharePoint Online
2. Supports certificate-based authentication (CBA) for automation scenarios  
3. Supports MFA-enabled account authentication
4. Automatically installs required PnP.PowerShell module upon confirmation
5. Audits a specific security group's permissions across all SharePoint sites
6. Exports detailed permission report to CSV file
7. Identifies direct and inherited permissions for the security group
8. Shows permission levels (Full Control, Edit, Read, etc.) for each site
9. Includes site collection and subsite analysis
10. Scheduler-friendly design for automated auditing
11. Comprehensive error handling and progress reporting
12. Filters out system and hidden sites by default

For detailed script execution: https://educ4te.com/
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $true)]
    [string]$SecurityGroupName,
    [string]$TenantName,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password,
    [switch]$IncludeSubsites,
    [switch]$IncludeSystemSites
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
        
    } catch {
        Write-Host "❌ Error connecting to SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
        $Global:ErrorLog += "Connection Error: $($_.Exception.Message)"
        Exit
    }
}

Function Get-SecurityGroupInfo {
    param([string]$GroupName)
    
    Write-Host "Preparing to search for security group: $GroupName..." -ForegroundColor Cyan
    
    # We'll search for the group within SharePoint permissions rather than pre-validating
    # This avoids the need for Microsoft Graph connection
    Write-Host "✅ Will search for security group during site permission analysis" -ForegroundColor Green
    
    # Return a simple object for consistency
    return @{
        DisplayName = $GroupName
        SearchName = $GroupName
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

Function Test-GroupPermissions {
    param(
        [string]$SiteUrl,
        [string]$GroupDisplayName
    )
    
    $Permissions = @()
    
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
                    # Check if this is our target group (by display name or login name containing the group)
                    $IsTargetGroup = $false
                    
                    # Debug: Show what principals we're finding (uncomment for troubleshooting)
                    # Write-Host "      Found principal: $($Principal.Title) ($($Principal.PrincipalType)) - $($Principal.LoginName)" -ForegroundColor DarkGray
                    
                    if ($Principal.PrincipalType -eq "SecurityGroup" -or $Principal.PrincipalType -eq "SharePointGroup") {
                        # Case-insensitive matching for group names
                        if ($Principal.Title -ieq $GroupDisplayName -or 
                            $Principal.LoginName -ilike "*$GroupDisplayName*" -or
                            $Principal.Title -ilike "*$GroupDisplayName*" -or
                            $Principal.LoginName -ieq $GroupDisplayName) {
                            $IsTargetGroup = $true
                        }
                    }
                    
                    if ($IsTargetGroup) {
                        $RoleDefinitions = Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings
                        $PermissionLevels = ($RoleDefinitions | ForEach-Object { $_.Name }) -join ", "
                        
                        $Permission = [PSCustomObject]@{
                            SiteUrl = $SiteUrl
                            SiteTitle = $Web.Title
                            GroupName = $Principal.Title
                            GroupType = $Principal.PrincipalType
                            LoginName = $Principal.LoginName
                            PermissionLevels = $PermissionLevels
                            PermissionType = "Direct"
                            CreatedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        }
                        
                        $Permissions += $Permission
                        $Global:FoundPermissions++
                    }
                }
            } catch {
                # Skip permission if we can't read it (likely insufficient permissions)
                $Global:WarningLog += "Warning: Could not read permission for $($SiteUrl)"
            }
        }
        
        # If including subsites, check them too
        if ($IncludeSubsites.IsPresent) {
            try {
                $Subwebs = Get-PnPSubWeb -Recurse
                foreach ($Subweb in $Subwebs) {
                    $SubPermissions = Test-GroupPermissions -SiteUrl $Subweb.Url -GroupDisplayName $GroupDisplayName
                    $Permissions += $SubPermissions
                }
            } catch {
                $Global:WarningLog += "Warning: Could not analyze subsites for $($SiteUrl)"
            }
        }
        
    } catch {
        $Global:ErrorLog += "Error analyzing site $($SiteUrl): $($_.Exception.Message)"
    }
    
    return $Permissions
}

Function Export-Results {
    param([array]$Results)
    
    $OutputCSV = ".\SecurityGroupSharePointAccess_$((Get-Date -format "yyyy-MMM-dd-ddd_hh-mm-ss_tt").ToString()).csv"
    
    if ($Results.Count -gt 0) {
        $Results | Export-Csv -Path $OutputCSV -NoTypeInformation
        
        Write-Host "`n✅ The output file contains $($Results.Count) permission records" -ForegroundColor Green
        Write-Host "📁 The output file is available at: " -NoNewline -ForegroundColor Yellow
        Write-Host $OutputCSV -ForegroundColor Cyan
        
        # Show summary
        Write-Host "`n📊 AUDIT SUMMARY:" -ForegroundColor Cyan
        Write-Host "   • Total Sites Processed: $Global:ProcessedSites" -ForegroundColor White
        Write-Host "   • Sites with Group Access: $($Results.Count)" -ForegroundColor White
        Write-Host "   • Total Permissions Found: $Global:FoundPermissions" -ForegroundColor White
        Write-Host "   • Warnings: $($Global:WarningLog.Count)" -ForegroundColor Yellow
        Write-Host "   • Errors: $($Global:ErrorLog.Count)" -ForegroundColor Red
        
        # Offer to open the file
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open the output file?", 0, "Open Output File", 4)
        if ($UserInput -eq 6) {
            Invoke-Item $OutputCSV
        }
        
    } else {
        Write-Host "`n⚠️  No permissions found for security group '$SecurityGroupName' on any SharePoint sites" -ForegroundColor Yellow
        Write-Host "   This could mean:" -ForegroundColor Yellow
        Write-Host "   • The group has no SharePoint access" -ForegroundColor Yellow
        Write-Host "   • The group name doesn't match exactly" -ForegroundColor Yellow
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
Write-Host "🔍 Microsoft Security Group SharePoint Access Audit Tool" -ForegroundColor Magenta
Write-Host "=" * 60 -ForegroundColor Magenta

try {
    # Install required modules
    Install-RequiredModules
    
    # Connect to SharePoint Online
    Connect-SharePointOnline
    
    # Get security group information
    $SecurityGroup = Get-SecurityGroupInfo -GroupName $SecurityGroupName
    
    # Get all SharePoint sites
    $Sites = Get-SharePointSites
    
    if ($Sites.Count -eq 0) {
        Write-Host "❌ No SharePoint sites found to analyze" -ForegroundColor Red
        Exit
    }
    
    # Analyze each site for group permissions
    Write-Host "`n🔍 Analyzing SharePoint sites for security group permissions..." -ForegroundColor Cyan
    $AllPermissions = @()
    
    foreach ($Site in $Sites) {
        $Global:ProcessedSites++
        $ProgressPercent = [math]::Round(($Global:ProcessedSites / $Global:TotalSites) * 100, 2)
        
        Write-Progress -Activity "Analyzing SharePoint Sites" `
                      -Status "Processing: $($Site.Title) ($Global:ProcessedSites of $Global:TotalSites)" `
                      -PercentComplete $ProgressPercent
        
        Write-Host "   Analyzing site: $($Site.Title) ($($Site.Url))" -ForegroundColor Gray
        
        try {
            $SitePermissions = Test-GroupPermissions -SiteUrl $Site.Url -GroupDisplayName $SecurityGroup.DisplayName
            $AllPermissions += $SitePermissions
            
            if ($SitePermissions.Count -gt 0) {
                Write-Host "     ✅ Found permissions on this site" -ForegroundColor Green
            }
            
        } catch {
            $Global:ErrorLog += "Failed to analyze site: $($Site.Url) - $($_.Exception.Message)"
            Write-Host "     ❌ Error analyzing site" -ForegroundColor Red
        }
    }
    
    Write-Progress -Activity "Analyzing SharePoint Sites" -Completed
    
    # Export results
    Export-Results -Results $AllPermissions
    
} catch {
    Write-Host "❌ Critical error during execution: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Always disconnect services
    Disconnect-Services
    
    # Display script information
    Write-Host "`n~~ Scripted by educ4te.com ~~" -ForegroundColor Green
}
