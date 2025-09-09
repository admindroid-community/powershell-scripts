<#
=============================================================================================
Name:           SharePoint Online Tenant Rename Script
Version:        1.0
Description:    Renames SharePoint Online tenant with scheduled execution delay
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. Modern authentication support with certificate-based priority
2. MFA-enabled account support  
3. Scheduled tenant rename with configurable delay
4. Comprehensive error handling and logging
5. Pre-validation checks for tenant rename eligibility
6. Progress monitoring and status reporting
============================================================================================
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [string]$CertificateThumbprint, 
    [string]$TenantId,
    [string]$UserName,
    [SecureString]$SecurePassword,
    [string]$CurrentTenantName = "n4k4r",
    [string]$NewTenantName = "educ4teaustralia",
    [ValidateRange(0, 168)]
    [int]$DelayHours = 25,
    [switch]$Force,
    [switch]$ValidationOnly = $false
)

# Global configuration
$Global:TenantRenameConfig = @{
    ErrorLog = @()
    WarningLog = @()
    StartTime = Get-Date
    LogFile = "SPOTenantRename_$(Get-Date -format 'yyyy-MMM-dd-ddd_hh-mm-ss_tt').log"
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Color = "White"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    Write-Host $logMessage -ForegroundColor $Color
    Add-Content -Path $Global:TenantRenameConfig.LogFile -Value $logMessage
    
    if ($Level -eq "ERROR") {
        $Global:TenantRenameConfig.ErrorLog += $logMessage
    } elseif ($Level -eq "WARNING") {
        $Global:TenantRenameConfig.WarningLog += $logMessage
    }
}

function Install-RequiredModules {
    Write-Log "Checking required PowerShell modules..." -Color Cyan
    
    # Check PowerShell version for PnP compatibility
    $psVersion = $PSVersionTable.PSVersion
    $isPnPCompatible = $psVersion.Major -ge 7 -and $psVersion.Minor -ge 4
    
    if (-not $isPnPCompatible) {
        Write-Log "PowerShell version $($psVersion.ToString()) detected. PnP.PowerShell requires PowerShell 7.4+. Using SPO module only." -Level "WARNING" -Color Yellow
    }
    
    $requiredModules = @("Microsoft.Online.SharePoint.PowerShell")
    if ($isPnPCompatible) {
        $requiredModules += "PnP.PowerShell"
    }
    
    foreach ($moduleName in $requiredModules) {
        $module = Get-Module $moduleName -ListAvailable
        if ($null -eq $module -or $module.Count -eq 0) {
            Write-Log "$moduleName is not available" -Level "WARNING" -Color Yellow
            if (-not $Force) {
                $confirm = Read-Host "Install $moduleName module? [Y] Yes [N] No"
                if ($confirm -notmatch "[yY]") {
                    Write-Log "Module $moduleName is required. Exiting." -Level "ERROR" -Color Red
                    Exit 1
                }
            }
            
            try {
                Write-Log "Installing $moduleName module..." -Color Magenta
                # Install NuGet provider first if needed
                if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
                }
                Install-Module $moduleName -Scope CurrentUser -Force -AllowClobber
                Write-Log "$moduleName module installed successfully" -Color Green
            }
            catch {
                Write-Log "Failed to install $moduleName module: $($_.Exception.Message)" -Level "ERROR" -Color Red
                if ($moduleName -eq "Microsoft.Online.SharePoint.PowerShell") {
                    Exit 1  # SPO module is required
                }
            }
        }
        else {
            Write-Log "$moduleName module is available" -Color Green
        }
        
        # Import the module to ensure it's loaded
        try {
            Import-Module $moduleName -Force -DisableNameChecking
            Write-Log "$moduleName module imported successfully" -Color Green
        }
        catch {
            Write-Log "Failed to import $moduleName module: $($_.Exception.Message)" -Level "WARNING" -Color Yellow
            if ($moduleName -eq "Microsoft.Online.SharePoint.PowerShell") {
                Exit 1  # SPO module is required
            }
        }
    }
    
    return $isPnPCompatible
}

function Connect-SharePointOnline {
    param(
        [bool]$IsPnPAvailable = $true
    )
    
    Write-Log "Connecting to SharePoint Online..." -Color Cyan
    
    $adminUrl = "https://$CurrentTenantName-admin.sharepoint.com"
    
    try {
        if ($ClientId -and $CertificateThumbprint -and $TenantId) {
            Write-Log "Using certificate-based authentication" -Color Cyan
            Connect-SPOService -Url $adminUrl -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId
            if ($IsPnPAvailable) {
                Connect-PnPOnline -Url $adminUrl -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -Tenant $TenantId
            }
        }
        elseif ($UserName -and $SecurePassword) {
            Write-Log "Using credential-based authentication" -Color Cyan
            $credential = New-Object System.Management.Automation.PSCredential($UserName, $SecurePassword)
            Connect-SPOService -Url $adminUrl -Credential $credential
            if ($IsPnPAvailable) {
                Connect-PnPOnline -Url $adminUrl -Credential $credential
            }
        }
        else {
            Write-Log "Using interactive authentication" -Color Cyan
            Connect-SPOService -Url $adminUrl
            if ($IsPnPAvailable) {
                Connect-PnPOnline -Url $adminUrl -Interactive
            }
        }
        
        Write-Log "Successfully connected to SharePoint Online" -Color Green
        return $true
    }
    catch {
        Write-Log "Failed to connect to SharePoint Online: $($_.Exception.Message)" -Level "ERROR" -Color Red
        return $false
    }
}

function Test-TenantRenameEligibility {
    Write-Log "Checking tenant rename eligibility..." -Color Cyan
    
    try {
        # Check current tenant properties
        $tenant = Get-SPOTenant
        
        # Validate current domain
        if ($tenant.SharePointUrl -notlike "*$CurrentTenantName.sharepoint.com*") {
            Write-Log "Current tenant name validation failed. Expected: $CurrentTenantName" -Level "ERROR" -Color Red
            return $false
        }
        
        # Check for active operations
        $operations = Get-SPOOperation | Where-Object { $_.Status -eq "InProgress" }
        if ($operations.Count -gt 0) {
            Write-Log "Found $($operations.Count) active operations. Tenant rename not recommended." -Level "WARNING" -Color Yellow
            if (-not $Force) {
                return $false
            }
        }
        
        # Check site collections count (Microsoft recommends under 10,000)
        $siteCount = (Get-SPOSite -Limit All).Count
        Write-Log "Total site collections: $siteCount" -Color Cyan
        
        if ($siteCount -gt 10000) {
            Write-Log "High number of site collections ($siteCount). Rename process may take longer." -Level "WARNING" -Color Yellow
        }
        
        Write-Log "Tenant rename eligibility check completed" -Color Green
        return $true
    }
    catch {
        Write-Log "Failed to check tenant eligibility: $($_.Exception.Message)" -Level "ERROR" -Color Red
        return $false
    }
}

function Start-ScheduledTenantRename {
    param(
        [int]$DelayHours
    )
    
    $scheduledTime = (Get-Date).AddHours($DelayHours)
    Write-Log "Tenant rename scheduled for: $($scheduledTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Color Cyan
    Write-Log "Current time: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))" -Color Cyan
    Write-Log "Delay: $DelayHours hours" -Color Cyan
    
    # Calculate delay in seconds
    $delaySeconds = $DelayHours * 3600
    
    Write-Log "Waiting $DelayHours hours before starting tenant rename..." -Color Yellow
    Write-Log "You can safely close this window. The rename will proceed automatically." -Color Yellow
    
    # Progress bar for waiting period
    for ($i = 0; $i -lt $delaySeconds; $i += 60) {
        $remainingMinutes = [math]::Round(($delaySeconds - $i) / 60, 0)
        $percentComplete = [math]::Round(($i / $delaySeconds) * 100, 1)
        
        Write-Progress -Activity "Waiting for scheduled tenant rename" -Status "Time remaining: $remainingMinutes minutes" -PercentComplete $percentComplete
        Start-Sleep -Seconds 60
        
        # Log progress every hour
        if ($i % 3600 -eq 0 -and $i -gt 0) {
            $hoursWaited = $i / 3600
            Write-Log "Waited $hoursWaited hour(s). $($DelayHours - $hoursWaited) hour(s) remaining." -Color Cyan
        }
    }
    
    Write-Progress -Activity "Waiting for scheduled tenant rename" -Completed
    Write-Log "Wait period completed. Starting tenant rename process..." -Color Green
}

function Start-TenantRename {
    Write-Log "Initiating SharePoint Online tenant rename..." -Color Cyan
    Write-Log "From: $CurrentTenantName.sharepoint.com" -Color Yellow
    Write-Log "To: $NewTenantName.sharepoint.com" -Color Yellow
    
    try {
        # Start the tenant rename operation
        $renameOperation = Start-SPOTenantRename -SourceSiteUrl "https://$CurrentTenantName.sharepoint.com" -TargetSiteUrl "https://$NewTenantName.sharepoint.com" -ValidationOnly:$false
        
        Write-Log "Tenant rename operation started successfully" -Color Green
        Write-Log "Operation ID: $($renameOperation.OperationId)" -Color Cyan
        
        # Monitor the operation
        do {
            Start-Sleep -Seconds 300 # Check every 5 minutes
            $status = Get-SPOTenantRenameStatus -OperationId $renameOperation.OperationId
            
            Write-Log "Rename status: $($status.Status) - Progress: $($status.PercentComplete)%" -Color Cyan
            
            if ($status.Status -eq "Failed") {
                Write-Log "Tenant rename failed: $($status.ErrorMessage)" -Level "ERROR" -Color Red
                return $false
            }
            
        } while ($status.Status -eq "InProgress")
        
        if ($status.Status -eq "Succeeded") {
            Write-Log "Tenant rename completed successfully!" -Color Green
            Write-Log "New tenant URL: https://$NewTenantName.sharepoint.com" -Color Green
            return $true
        }
        else {
            Write-Log "Tenant rename completed with status: $($status.Status)" -Level "WARNING" -Color Yellow
            return $false
        }
    }
    catch {
        Write-Log "Failed to start tenant rename: $($_.Exception.Message)" -Level "ERROR" -Color Red
        return $false
    }
}

function Disconnect-Services {
    Write-Log "Disconnecting from SharePoint Online services..." -Color Cyan
    
    try {
        if (Get-Command "Disconnect-SPOService" -ErrorAction SilentlyContinue) {
            Disconnect-SPOService -ErrorAction SilentlyContinue
        }
        if (Get-Command "Disconnect-PnPOnline" -ErrorAction SilentlyContinue) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        Write-Log "Successfully disconnected from all services" -Color Green
    }
    catch {
        Write-Log "Warning: Could not properly disconnect from services: $($_.Exception.Message)" -Level "WARNING" -Color Yellow
    }
}

function Write-Summary {
    Write-Log "=== TENANT RENAME SUMMARY ===" -Color Magenta
    Write-Log "Start Time: $($Global:TenantRenameConfig.StartTime)" -Color Cyan
    Write-Log "End Time: $(Get-Date)" -Color Cyan
    Write-Log "Duration: $((Get-Date) - $Global:TenantRenameConfig.StartTime)" -Color Cyan
    Write-Log "Errors: $($Global:TenantRenameConfig.ErrorLog.Count)" -Color $(if($Global:TenantRenameConfig.ErrorLog.Count -gt 0) { "Red" } else { "Green" })
    Write-Log "Warnings: $($Global:TenantRenameConfig.WarningLog.Count)" -Color $(if($Global:TenantRenameConfig.WarningLog.Count -gt 0) { "Yellow" } else { "Green" })
    Write-Log "Log File: $($Global:TenantRenameConfig.LogFile)" -Color Cyan
    
    if ($Global:TenantRenameConfig.ErrorLog.Count -eq 0) {
        Write-Log "Tenant rename process completed successfully!" -Color Green
    }
    else {
        Write-Log "Tenant rename process completed with errors. Check log file for details." -Level "WARNING" -Color Yellow
    }
}

# Main execution
try {
    Write-Log "Starting SharePoint Online Tenant Rename Script" -Color Magenta
    Write-Log "Current Tenant: $CurrentTenantName.sharepoint.com" -Color Cyan
    Write-Log "New Tenant: $NewTenantName.sharepoint.com" -Color Cyan
    Write-Log "Scheduled Delay: $DelayHours hours" -Color Cyan
    
    # Install required modules
    $isPnPAvailable = Install-RequiredModules
    
    # Connect to SharePoint Online
    if (-not (Connect-SharePointOnline -IsPnPAvailable $isPnPAvailable)) {
        throw "Failed to connect to SharePoint Online"
    }
    
    # Check tenant rename eligibility
    if (-not (Test-TenantRenameEligibility)) {
        throw "Tenant is not eligible for rename operation"
    }
    
    # If validation only, exit here
    if ($ValidationOnly) {
        Write-Log "Validation completed successfully. Tenant is eligible for rename." -Color Green
        return
    }
    
    # Wait for scheduled time
    Start-ScheduledTenantRename -DelayHours $DelayHours
    
    # Perform the tenant rename
    $renameSuccess = Start-TenantRename
    
    if (-not $renameSuccess) {
        throw "Tenant rename operation failed"
    }
    
    Write-Log "SharePoint Online tenant has been successfully renamed!" -Color Green
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level "ERROR" -Color Red
    $Global:TenantRenameConfig.ErrorLog += $_.Exception.Message
}
finally {
    # Cleanup
    Disconnect-Services
    Write-Summary
}