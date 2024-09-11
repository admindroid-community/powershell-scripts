<#
=============================================================================================

Name         : Get List Item Count in SharePoint Online Using PowerShell  
Version      : 1.0
website      : o365reports.com

-----------------
Script Highlights
-----------------
1. The script exports SharePoint Online lists and their items count in the organization. 
2. Provides the SPO list item count of a single site. 
3. Allows to retrieve SPO list items count of multiple sites using an input CSV.  
4. Helps to get the total list items count including hidden lists.  
5. Tracks the inactive lists and their item count in SharePoint.  
6. Automatically installs the PnP PowerShell module (if not installed already) upon your confirmation. 
7. The script can be executed with an MFA-enabled account too. 
8. Exports report results as a CSV file. 
9. The script is schedular-friendly. 
10. The script uses modern authentication to connect SharePoint Online. 
11. It can be executed with certificate-based authentication (CBA) too. 

For detailed Script execution:  https://o365reports.com/2024/09/03/get-list-item-count-in-sharepoint-online-using-powershell/
============================================================================================
\#>param
(
    [Parameter(Mandatory = $false)]
    [string]$AdminName,
    [string]$Password,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$SiteUrl,
    [string]$ImportCsv,
    [string]$TenantName,
    [switch]$IncludeHiddenLists,
    [switch]$ShowHiddenListsOnly,
    [int]$Inactivedays
)

$Global:Tenant = $TenantName

# PnP Installation module
Function Installation-Module {
    $Module = Get-InstalledModule -Name PnP.PowerShell -ErrorAction SilentlyContinue
    if ($Module.Count -eq 0) {
        Write-Host "PnP PowerShell Module is not available." -ForegroundColor Yellow
        $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-Host "Installing PnP PowerShell module..."
            Install-Module PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
            Import-Module -Name PnP.PowerShell -Force
            #Register a new Azure AD Application and Grant Access to the tenant
            Register-PnPManagementShellAccess
        } else {
            Write-Host "PnP PowerShell module is required to connect to SharePoint Online. Please install the module using the Install-Module PnP.PowerShell cmdlet."
            Exit
        }
    }
    Write-Host "`nConnecting to SharePoint Online..."
}

# Connection implementation module
Function Connection-Module {
    param(
        [string]$Url
    )
    if (($AdminName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential = New-Object System.Management.Automation.PSCredential $AdminName, $SecuredPassword
        Connect-PnPOnline -Url $Url -Credential $Credential -WarningAction SilentlyContinue
    }
    elseif ($ClientId -ne "" -and $CertificateThumbprint -ne "" -and $TenantName -ne "") {
        Connect-PnPOnline -Url $Url -ClientId $ClientId -Thumbprint $CertificateThumbprint -Tenant "$TenantName.onmicrosoft.com"
    }
    else {
        Connect-PnPOnline -Url $Url -Interactive -WarningAction SilentlyContinue
    }
}

# Main function which exports SharePoint Online lists and its items count
Function Export-ListItemCounts {

    if ($SiteUrl -ne "") {
        $SiteCollections = $SiteUrl
    }
    elseif ($ImportCsv -ne "") {
        $SiteURLs = Import-Csv -Path $ImportCsv
        $SiteCollections = $SiteURLs.SiteUrl
    }
    else {
        if ($TenantName -eq "") {
            $Global:Tenant = $(Write-Host "`nEnter your tenant name (e.g., contoso): " -ForegroundColor Green -NoNewline; Read-Host)
        }
        $SPO_admin_url = "https://$Global:Tenant-admin.sharepoint.com"
        Connection-Module -Url $SPO_admin_url
        $SiteCollections = Get-PnPTenantSite | Select-Object -ExpandProperty Url
    }

    $FilteredSiteCollections = $SiteCollections | Where-Object { $_ -notlike "*-my.sharepoint.com*" }

    $ListInventory = @()
    $CurrentSite = 0
    $TotalSites = $FilteredSiteCollections.Count
    $Global:AccessDeniedCount = 0

    
    foreach ($Site in $FilteredSiteCollections) {
        $CurrentSite++
        Write-Progress -Activity "Processing $Site" -Status "Processing $CurrentSite out of $TotalSites sites" -PercentComplete (($CurrentSite / $TotalSites) * 100)

        Try {
            
            Connection-Module -Url $Site

            # Get all subsites (including root site)
            $SubWebs = Get-PnPSubWeb -Recurse -IncludeRootWeb

            foreach ($Web in $SubWebs) {

                Connection-Module -Url $Web.Url

                $ExcludedLists = @("Reusable Content", "Content and Structure Reports", "Form Templates", "Images", "Pages", "Workflow History", "Workflow Tasks", "Preservation Hold Library", "Tenant Wide Extensions", "Hub Settings", "Channel Settings")

                # Gets each site list details
                Get-PnPList -Includes Author | Where-Object {
                    ($ShowHiddenListsOnly.IsPresent -and $_.Hidden) -or (-not $ShowHiddenListsOnly.IsPresent -and ($IncludeHiddenLists.IsPresent  -or -not $_.Hidden)) -and ($ExcludedLists -notcontains $_.Title) -and ($_.BaseType -ne "DocumentLibrary")
                } | ForEach-Object {
                    if ($_.RootFolder.ServerRelativeUrl -like "*/Lists/*") {
                        $Data = @{
                            'Created Time' = $_.Created
                            'Author Name' = $_.Author.Title
                            'Author Email' = $_.Author.Email
                            'List Title' = $_.Title
                            'Is Hidden' = $_.Hidden
                            'List URL' = $_.RootFolder.ServerRelativeUrl
                            'Item Count' = $_.ItemCount
                            'Default View URL' = $_.DefaultViewUrl
                            'Site Name' = $Web.Title
                            'Site URL' = $Web.Url
                            'Last Item Deleted Date' = $_.LastItemDeletedDate
                            'Last Item User Modified Date' = $_.LastItemUserModifiedDate
                            'Inactive Days' = (New-TimeSpan -Start $_.LastItemUserModifiedDate).Days
                            'Is Attachment Enabled' = $_.EnableAttachments
                        }

                        # Inactive lists checking section
                        if (-not $Inactivedays -or ($Data['Inactive Days'] -gt $Inactivedays)) {
                            $ListInventory = New-Object PSObject -Property $Data
                            $ListInventory | Select-Object 'List Title', 'List URL', 'Item Count', 'Site Name', 'Last Item User Modified Date', 'Inactive Days', 'Author Name', 'Author Email', 'Created Time', 'Last Item Deleted Date', 'Is Hidden', 'Is Attachment Enabled', 'Default View URL', 'Site URL' | Export-CSV -Path $ReportOutput -NoTypeInformation -Append -Force
                        }
                    }
                }
            }
        }
        Catch {
            if ($_.Exception.Message) {
                $Global:AccessDeniedCount++
            }
        }
    }
}

Installation-Module
$ReportOutput = "$PSScriptRoot\SPO_List_Item_Count_Report_$((Get-Date -Format yyyy-MM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
Export-ListItemCounts
# Disconnects PnPOnline session
Disconnect-PnPOnline

if ((Test-Path -Path $ReportOutput) -eq "True") {
    if ($Global:AccessDeniedCount -gt 0) {
        Write-Host "`nYou don't have access on $Global:AccessDeniedCount site(s)." -ForegroundColor Yellow
    }
    Write-Host "`nThe Output file is available in: " -NoNewline -ForegroundColor Yellow
    Write-Host $ReportOutput
    Write-Host "`n~~ Script prepared by AdminDroid Community ~~`n" -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green "`n`n"
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open the output file?", 0, "Open Output File", 4)
    If ($UserInput -eq 6) {
        Invoke-Item "$ReportOutput"
    }
} else {
    Write-Host -ForegroundColor Yellow "No Records Found"
}