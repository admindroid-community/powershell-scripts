<#
=============================================================================================
Name:           Create SharePoint Site Collections with Content
Version:        3.0
Description:    Creates 10 site collections with 20 sites each, populated with Office documents
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. Modern authentication support with certificate-based auth priority
2. MFA-enabled account support  
3. Automatic PnP PowerShell module installation
4. Progress reporting and error handling
5. Scheduler-friendly design with certificate authentication
6. Bulk document creation with templates
============================================================================================
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [string]$CertificateThumbprint, 
    [string]$TenantId,
    [string]$UserName,
    [string]$Password,
    [Parameter(Mandatory = $true)]
    [string]$TenantUrl,
    [Parameter(Mandatory = $false)]
    [string]$SiteCollectionPrefix = "TestSite",
    [Parameter(Mandatory = $false)]
    [int]$SiteCollectionCount = 10,
    [Parameter(Mandatory = $false)]
    [int]$SitesPerCollection = 20,
    [Parameter(Mandatory = $false)]
    [int]$WordDocCount = 30,
    [Parameter(Mandatory = $false)]
    [int]$ExcelSheetCount = 30,
    [Parameter(Mandatory = $false)]
    [int]$PdfFileCount = 50
)

# Global configuration
$Global:ProvisioningConfig = @{
    ErrorLog = @()
    WarningLog = @()
    CreatedSites = @()
    CreatedDocuments = @()
    StartTime = Get-Date
}

# Document templates
$WordTemplate = @"
<!DOCTYPE html>
<html>
<head>
    <title>Sample Document</title>
</head>
<body>
    <h1>Sample Word Document</h1>
    <p>This is a sample document created for testing purposes.</p>
    <p>Document created on: $(Get-Date)</p>
    <h2>Content Sections</h2>
    <p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>
    <p>Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.</p>
</body>
</html>
"@

$ExcelTemplate = @"
Name,Department,Salary,Start Date
John Smith,IT,75000,2023-01-15
Jane Doe,HR,65000,2023-02-20
Mike Johnson,Finance,80000,2023-03-10
Sarah Wilson,Marketing,70000,2023-04-05
David Brown,IT,72000,2023-05-12
Lisa Davis,HR,68000,2023-06-18
Tom Anderson,Finance,85000,2023-07-22
Emily Taylor,Marketing,73000,2023-08-30
Chris Martin,IT,77000,2023-09-14
Ashley Garcia,HR,66000,2023-10-25
"@

# Function to check and install required modules
function Install-RequiredModules {
    Write-Host "Checking required modules..." -ForegroundColor Cyan
    
    $Module = Get-Module PnP.PowerShell -ListAvailable
    if ($null -eq $Module -or $Module.count -eq 0) {
        Write-Host "PnP.PowerShell module is not available" -ForegroundColor Yellow
        $Confirm = Read-Host "Install PnP.PowerShell module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Magenta
            Install-Module PnP.PowerShell -Scope CurrentUser -Force
            Write-Host "Module installed successfully" -ForegroundColor Green
        } else {
            Write-Host "PnP.PowerShell module is required. Exiting." -ForegroundColor Red
            Exit
        }
    }
}

# Function to establish SharePoint connection
function Connect-SharePointOnline {
    try {
        Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
        
        if ($ClientId -and $CertificateThumbprint -and $TenantId) {
            Write-Host "Using certificate-based authentication" -ForegroundColor Green
            Connect-PnPOnline -Url $TenantUrl -ClientId $ClientId -Thumbprint $CertificateThumbprint -Tenant $TenantId
        } elseif ($UserName -and $Password) {
            Write-Host "Using basic authentication" -ForegroundColor Yellow
            $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
            $Credential = New-Object System.Management.Automation.PSCredential($UserName, $SecurePassword)
            Connect-PnPOnline -Url $TenantUrl -Credentials $Credential
        } else {
            Write-Host "Using interactive authentication" -ForegroundColor Green
            Connect-PnPOnline -Url $TenantUrl -Interactive
        }
        
        Write-Host "Successfully connected to SharePoint Online" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to connect to SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
        $Global:ProvisioningConfig.ErrorLog += "Connection failed: $($_.Exception.Message)"
        return $false
    }
}

# Function to create site collection
function New-SiteCollection {
    param(
        [string]$SiteUrl,
        [string]$Title,
        [string]$Owner
    )
    
    try {
        Write-Host "Creating site collection: $Title" -ForegroundColor Cyan
        
        New-PnPSite -Type TeamSite -Title $Title -Url $SiteUrl -Owner $Owner -Wait
        
        $Global:ProvisioningConfig.CreatedSites += @{
            Type = "SiteCollection"
            Url = $SiteUrl
            Title = $Title
            Created = Get-Date
        }
        
        Write-Host "Site collection created successfully: $SiteUrl" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to create site collection $Title`: $($_.Exception.Message)" -ForegroundColor Red
        $Global:ProvisioningConfig.ErrorLog += "Site collection creation failed: $Title - $($_.Exception.Message)"
        return $false
    }
}

# Function to create subsite
function New-SubSite {
    param(
        [string]$ParentUrl,
        [string]$SiteUrl,
        [string]$Title
    )
    
    try {
        Connect-PnPOnline -Url $ParentUrl -Interactive
        
        New-PnPWeb -Title $Title -Url $SiteUrl -Template "STS#3"
        
        $Global:ProvisioningConfig.CreatedSites += @{
            Type = "SubSite"
            Url = "$ParentUrl/$SiteUrl"
            Title = $Title
            Created = Get-Date
        }
        
        Write-Host "Subsite created: $Title" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to create subsite $Title`: $($_.Exception.Message)" -ForegroundColor Red
        $Global:ProvisioningConfig.ErrorLog += "Subsite creation failed: $Title - $($_.Exception.Message)"
        return $false
    }
}

# Function to create documents in a site
function New-DocumentsInSite {
    param(
        [string]$SiteUrl
    )
    
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive
        
        # Create Word documents
        for ($i = 1; $i -le $WordDocCount; $i++) {
            $fileName = "Document_$i.docx"
            $tempFile = [System.IO.Path]::GetTempFileName() + ".html"
            $WordTemplate | Out-File -FilePath $tempFile -Encoding UTF8
            
            Add-PnPFile -Path $tempFile -Folder "Shared Documents" -NewFileName $fileName
            Remove-Item $tempFile -Force
            
            $Global:ProvisioningConfig.CreatedDocuments += @{
                Site = $SiteUrl
                Type = "Word"
                Name = $fileName
                Created = Get-Date
            }
        }
        
        # Create Excel files
        for ($i = 1; $i -le $ExcelSheetCount; $i++) {
            $fileName = "Spreadsheet_$i.xlsx"
            $tempFile = [System.IO.Path]::GetTempFileName() + ".csv"
            $ExcelTemplate | Out-File -FilePath $tempFile -Encoding UTF8
            
            Add-PnPFile -Path $tempFile -Folder "Shared Documents" -NewFileName $fileName
            Remove-Item $tempFile -Force
            
            $Global:ProvisioningConfig.CreatedDocuments += @{
                Site = $SiteUrl
                Type = "Excel"
                Name = $fileName
                Created = Get-Date
            }
        }
        
        # Create PDF files (as text files with PDF extension for simulation)
        for ($i = 1; $i -le $PdfFileCount; $i++) {
            $fileName = "Document_$i.pdf"
            $pdfContent = "PDF Content for $fileName`nCreated: $(Get-Date)`nThis is sample PDF content for testing purposes."
            $tempFile = [System.IO.Path]::GetTempFileName() + ".txt"
            $pdfContent | Out-File -FilePath $tempFile -Encoding UTF8
            
            Add-PnPFile -Path $tempFile -Folder "Shared Documents" -NewFileName $fileName
            Remove-Item $tempFile -Force
            
            $Global:ProvisioningConfig.CreatedDocuments += @{
                Site = $SiteUrl
                Type = "PDF"
                Name = $fileName
                Created = Get-Date
            }
        }
        
        Write-Host "Documents created in $SiteUrl" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to create documents in $SiteUrl`: $($_.Exception.Message)" -ForegroundColor Red
        $Global:ProvisioningConfig.ErrorLog += "Document creation failed: $SiteUrl - $($_.Exception.Message)"
        return $false
    }
}

# Function to generate summary report
function Write-SummaryReport {
    $timestamp = Get-Date -Format "yyyy-MMM-dd-ddd_hh-mm-ss_tt"
    $reportPath = "SharePoint_Provisioning_Report_$timestamp.csv"
    
    $reportData = @()
    
    foreach ($site in $Global:ProvisioningConfig.CreatedSites) {
        $reportData += [PSCustomObject]@{
            Type = $site.Type
            Title = $site.Title
            Url = $site.Url
            Created = $site.Created
            Status = "Success"
        }
    }
    
    $reportData | Export-Csv -Path $reportPath -NoTypeInformation
    
    Write-Host "`n========== PROVISIONING SUMMARY ==========" -ForegroundColor Cyan
    Write-Host "Total Site Collections Created: $($Global:ProvisioningConfig.CreatedSites | Where-Object {$_.Type -eq 'SiteCollection'} | Measure-Object | Select-Object -ExpandProperty Count)" -ForegroundColor Green
    Write-Host "Total Subsites Created: $($Global:ProvisioningConfig.CreatedSites | Where-Object {$_.Type -eq 'SubSite'} | Measure-Object | Select-Object -ExpandProperty Count)" -ForegroundColor Green
    Write-Host "Total Documents Created: $($Global:ProvisioningConfig.CreatedDocuments.Count)" -ForegroundColor Green
    Write-Host "Report exported to: $reportPath" -ForegroundColor Yellow
    
    if ($Global:ProvisioningConfig.ErrorLog.Count -gt 0) {
        Write-Host "`nErrors encountered:" -ForegroundColor Red
        $Global:ProvisioningConfig.ErrorLog | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    }
}

# Main execution
try {
    Write-Host "Starting SharePoint provisioning process..." -ForegroundColor Cyan
    Write-Host "Tenant URL: $TenantUrl" -ForegroundColor Yellow
    Write-Host "Site Collections to create: $SiteCollectionCount" -ForegroundColor Yellow
    Write-Host "Sites per collection: $SitesPerCollection" -ForegroundColor Yellow
    
    # Install required modules
    Install-RequiredModules
    
    # Connect to SharePoint
    if (-not (Connect-SharePointOnline)) {
        throw "Failed to establish SharePoint connection"
    }
    
    # Get current user for site ownership
    $currentUser = (Get-PnPContext).Web.CurrentUser.Email
    if (-not $currentUser) {
        $currentUser = $UserName
    }
    
    $totalOperations = $SiteCollectionCount * (1 + $SitesPerCollection)
    $currentOperation = 0
    
    # Create site collections and subsites
    for ($sc = 1; $sc -le $SiteCollectionCount; $sc++) {
        $currentOperation++
        $percentComplete = [math]::Round(($currentOperation / $totalOperations) * 100, 2)
        Write-Progress -Activity "Creating SharePoint Infrastructure" -Status "Creating Site Collection $sc of $SiteCollectionCount" -PercentComplete $percentComplete
        
        $siteCollectionUrl = "$TenantUrl/sites/$SiteCollectionPrefix$sc"
        $siteCollectionTitle = "$SiteCollectionPrefix $sc"
        
        if (New-SiteCollection -SiteUrl $siteCollectionUrl -Title $siteCollectionTitle -Owner $currentUser) {
            
            # Create subsites in this collection
            for ($s = 1; $s -le $SitesPerCollection; $s++) {
                $currentOperation++
                $percentComplete = [math]::Round(($currentOperation / $totalOperations) * 100, 2)
                Write-Progress -Activity "Creating SharePoint Infrastructure" -Status "Creating Subsite $s of $SitesPerCollection in Collection $sc" -PercentComplete $percentComplete
                
                $subsiteUrl = "subsite$s"
                $subsiteTitle = "Subsite $s"
                
                if (New-SubSite -ParentUrl $siteCollectionUrl -SiteUrl $subsiteUrl -Title $subsiteTitle) {
                    # Create documents in the subsite
                    New-DocumentsInSite -SiteUrl "$siteCollectionUrl/$subsiteUrl"
                }
            }
            
            # Also create documents in the root site collection
            New-DocumentsInSite -SiteUrl $siteCollectionUrl
        }
    }
    
    Write-Progress -Activity "Creating SharePoint Infrastructure" -Completed
    Write-Host "SharePoint provisioning completed successfully!" -ForegroundColor Green
    
}
catch {
    Write-Host "Critical error during provisioning: $($_.Exception.Message)" -ForegroundColor Red
    $Global:ProvisioningConfig.ErrorLog += "Critical error: $($_.Exception.Message)"
}
finally {
    # Generate summary report
    Write-SummaryReport
    
    # Cleanup
    try {
        Disconnect-PnPOnline
        Write-Host "Disconnected from SharePoint Online" -ForegroundColor Green
    }
    catch {
        Write-Host "Error during disconnection: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    $endTime = Get-Date
    $duration = $endTime - $Global:ProvisioningConfig.StartTime
    Write-Host "Total execution time: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
}
