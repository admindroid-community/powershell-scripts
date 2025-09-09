#Requires -Version 7.0

<#
.SYNOPSIS
    Test Microsoft 365 connection and module compatibility
.DESCRIPTION
    Simple test script to verify Microsoft 365 modules and connections work properly
.VERSION
    1.0.0
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId = ""
)

# Test function to check module compatibility
function Test-M365Modules {
    Write-Host "🔍 Testing Microsoft 365 Modules..." -ForegroundColor Cyan
    
    $modules = @(
        @{ Name = "Microsoft.Graph"; MinVersion = "2.0.0" },
        @{ Name = "PnP.PowerShell"; MinVersion = "1.12.0" },
        @{ Name = "ExchangeOnlineManagement"; MinVersion = "3.0.0" }
    )
    
    $allGood = $true
    
    foreach ($module in $modules) {
        try {
            $installed = Get-Module -Name $module.Name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            
            if ($installed) {
                if ($installed.Version -ge [version]$module.MinVersion) {
                    Write-Host "✅ $($module.Name) v$($installed.Version) - OK" -ForegroundColor Green
                } else {
                    Write-Host "⚠️  $($module.Name) v$($installed.Version) - Outdated (need $($module.MinVersion)+)" -ForegroundColor Yellow
                    $allGood = $false
                }
            } else {
                Write-Host "❌ $($module.Name) - Not installed" -ForegroundColor Red
                $allGood = $false
            }
        }
        catch {
            Write-Host "❌ $($module.Name) - Error checking: $($_.Exception.Message)" -ForegroundColor Red
            $allGood = $false
        }
    }
    
    return $allGood
}

# Test Graph connection
function Test-GraphConnection {
    Write-Host "`n🔍 Testing Microsoft Graph Connection..." -ForegroundColor Cyan
    
    try {
        # Try basic connection first
        $scopes = @("User.Read", "User.Read.All")
        
        if ($TenantId) {
            Connect-MgGraph -Scopes $scopes -TenantId $TenantId -NoWelcome -ErrorAction Stop
        } else {
            Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
        }
        
        $context = Get-MgContext
        Write-Host "✅ Connected to Microsoft Graph" -ForegroundColor Green
        Write-Host "   Account: $($context.Account)" -ForegroundColor Gray
        Write-Host "   Tenant: $($context.TenantId)" -ForegroundColor Gray
        
        # Test basic query
        try {
            Get-MgUser -Top 1 -ErrorAction Stop | Out-Null
            Write-Host "✅ Basic user query successful" -ForegroundColor Green
        }
        catch {
            Write-Host "⚠️  Basic user query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # Test guest user query
        try {
            $guests = Get-MgUser -Filter "userType eq 'Guest'" -Top 5 -ErrorAction Stop
            Write-Host "✅ Guest user query successful ($($guests.Count) found)" -ForegroundColor Green
        }
        catch {
            Write-Host "⚠️  Guest user query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        return $true
    }
    catch {
        Write-Host "❌ Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Test SharePoint connection  
function Test-SharePointConnection {
    Write-Host "`n🔍 Testing SharePoint Connection..." -ForegroundColor Cyan
    
    try {
        # Try to determine admin URL
        $adminUrl = "https://m365x22747677-admin.sharepoint.com"
        
        try {
            $context = Get-MgContext
            if ($context) {
                $org = Get-MgOrganization -ErrorAction SilentlyContinue
                if ($org.VerifiedDomains) {
                    $primaryDomain = ($org.VerifiedDomains | Where-Object { $_.IsInitial -eq $true }).Name
                    if ($primaryDomain) {
                        $tenantName = $primaryDomain.Split('.')[0]
                        $adminUrl = "https://$tenantName-admin.sharepoint.com"
                    }
                }
            }
        }
        catch {
            Write-Host "Using fallback admin URL" -ForegroundColor Yellow
        }
        
        Write-Host "Attempting connection to: $adminUrl" -ForegroundColor Gray
        
        Connect-PnPOnline -Url $adminUrl -Interactive -ErrorAction Stop
        
        Write-Host "✅ Connected to SharePoint Online" -ForegroundColor Green
        
        # Test basic query
        try {
            $sites = Get-PnPTenantSite -Top 5 -ErrorAction Stop
            Write-Host "✅ Site collection query successful ($($sites.Count) found)" -ForegroundColor Green
        }
        catch {
            Write-Host "⚠️  Site collection query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        return $true
    }
    catch {
        Write-Host "❌ SharePoint connection failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Test Exchange Online connection
function Test-ExchangeConnection {
    Write-Host "`n🔍 Testing Exchange Online Connection..." -ForegroundColor Cyan
    
    try {
        # Try modern authentication
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        Write-Host "✅ Connected to Exchange Online" -ForegroundColor Green
        
        # Test basic query
        try {
            $mailboxes = Get-EXOMailbox -ResultSize 5 -ErrorAction Stop
            Write-Host "✅ Mailbox query successful ($($mailboxes.Count) found)" -ForegroundColor Green
        }
        catch {
            Write-Host "⚠️  Mailbox query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        return $true
    }
    catch {
        Write-Host "❌ Exchange Online connection failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "   This could be due to MSAL runtime issues" -ForegroundColor Yellow
        return $false
    }
}

# Main execution
try {
    Write-Host @"
╔══════════════════════════════════════════════════════════════╗
║             Microsoft 365 Connection Test Tool               ║
║                        Version 1.0.0                         ║
╚══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan

    # Test modules
    $modulesOK = Test-M365Modules
    
    if (-not $modulesOK) {
        Write-Host "`n❌ Module issues detected. Please install/update required modules before proceeding." -ForegroundColor Red
        exit 1
    }
    
    # Test connections
    $graphOK = Test-GraphConnection
    $spoOK = Test-SharePointConnection
    $exoOK = Test-ExchangeConnection
    
    # Summary
    Write-Host "`n" -NoNewline
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "                    TEST RESULTS                       " -ForegroundColor Green  
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "Microsoft Graph:     $(if($graphOK){'✅ PASS'}else{'❌ FAIL'})" -ForegroundColor $(if($graphOK){'Green'}else{'Red'})
    Write-Host "SharePoint Online:   $(if($spoOK){'✅ PASS'}else{'❌ FAIL'})" -ForegroundColor $(if($spoOK){'Green'}else{'Red'})
    Write-Host "Exchange Online:     $(if($exoOK){'✅ PASS'}else{'❌ FAIL'})" -ForegroundColor $(if($exoOK){'Green'}else{'Red'})
    
    if ($graphOK -and $spoOK) {
        Write-Host "`n✅ Core services are working. The audit script should run successfully." -ForegroundColor Green
        if (-not $exoOK) {
            Write-Host "⚠️  Exchange Online issues may limit some group analysis features." -ForegroundColor Yellow
        }
    } else {
        Write-Host "`n❌ Critical services failed. Please resolve connection issues before running the audit." -ForegroundColor Red
    }
}
catch {
    Write-Host "`n❌ Test failed with error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Cleanup connections
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Disconnect-PnPOnline -ErrorAction SilentlyContinue  
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore cleanup errors
    }
}
