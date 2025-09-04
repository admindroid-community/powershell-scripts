<#
.SYNOPSIS
    Test script for the fixed M365 External User Audit
    
.DESCRIPTION
    This script tests the fixed version with enhanced error handling and module management
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\SecureAudit\Reports\Test",
    
    [Parameter(Mandatory = $false)]
    [int]$DaysToAudit = 30
)

# Clear any previous errors
$Error.Clear()

try {
    Write-Host "🧪 Testing Fixed M365 External User Audit Script" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    Write-Host "Test Parameters:" -ForegroundColor Yellow
    Write-Host "  Output Path: $OutputPath" -ForegroundColor White
    Write-Host "  Days to Audit: $DaysToAudit" -ForegroundColor White
    Write-Host ""
    
    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-Host "✅ Created output directory: $OutputPath" -ForegroundColor Green
    }
    
    # Execute the fixed audit script
    $scriptPath = Join-Path $PSScriptRoot "M365-External-User-Audit-Fixed.ps1"
    
    if (-not (Test-Path $scriptPath)) {
        throw "Fixed audit script not found: $scriptPath"
    }
    
    Write-Host "🚀 Executing fixed audit script..." -ForegroundColor Green
    Write-Host ""
    
    # Run the script with minimal parameters for testing
    & $scriptPath -OutputPath $OutputPath -DaysToAudit $DaysToAudit
    
    Write-Host ""
    Write-Host "✅ Test execution completed!" -ForegroundColor Green
    Write-Host "Check the output directory for generated reports: $OutputPath" -ForegroundColor Cyan
}
catch {
    Write-Host ""
    Write-Host "❌ Test execution failed!" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.ScriptStackTrace) {
        Write-Host ""
        Write-Host "Stack Trace:" -ForegroundColor Yellow
        Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    }
    
    exit 1
}
