<#
=========================================================================================
Name:           Allow External Sharing for Specific SharePoint Sites
Version:        1.0
Website:        blog.admindroid.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script utilizes SharePoint PnP PowerShell and installs it (if not already installed) upon your confirmation.
2. The script allows you to configure external sharing for a single site.
3. The script supports configuring external sharing for multiple sites.
4. Admins can restrict external sharing for the remaining sites.
5. It helps register the Entra ID app needed to use PnP PowerShell.
6. The script exports a log file.
7. It can be executed with an MFA-enabled account.
8. This script also supports certificate-based authentication (CBA).
9. The script is scheduler-friendly.

For detailed script execution: https://blog.admindroid.com/allow-external-sharing-for-specific-sharepoint-sites/ 

=========================================================================================
#>
param
(
    [Parameter(Mandatory = $false)]
    [ValidateSet(
        'Only people in your org',
        'Existing guests',
        'New and existing guests',
        'Anyone'
    )]
    [string]$SharingConfigForRemainingSites,
    [string]$AdminName,
    [string]$Password,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$SiteUrl,
    [string]$ImportCsv,
    [string]$TenantName,
    [switch]$RegisterNewApp
)

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
            # Register a new Azure AD Application and Grant Access to the tenant
            Register-PnPManagementShellAccess
        } else {
            Write-Host "PnP PowerShell module is required to connect to SharePoint Online. Please install the module using the Install-Module PnP.PowerShell cmdlet."
            Exit
        }
    }
}

# PnP Interactive Sign-in App Registration Module
Function Interactive-App-Register {
    $AppRegistered = $false
    while (-not $AppRegistered) {
        if ($script:TenantName -eq "") {
                $script:TenantName = $(Write-Host "`nEnter your tenant name (e.g., contoso): " -NoNewline; Read-Host)
        }
        try {
            $AppName = $(Write-Host "`nEnter name for your application: " -NoNewline; Read-Host)
            $PnPApp = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName $AppName -Tenant "$script:TenantName.onmicrosoft.com" -Interactive -WarningAction SilentlyContinue
            $script:ClientId = $PnPApp."AzureAppId/ClientId"
            $AppRegistered = $true
            Write-Host "Save this ClientId for future use: $script:ClientId`n" -ForegroundColor Blue
        } catch {
            $_.Exception.Message
        }
    }
}

# Connection Module
Function Connection-Module {
    
    Write-Host "`nConnecting to SharePoint Online..."

    if ($script:TenantName -eq "") {
        $script:TenantName = Read-Host "Enter your tenant name (e.g., contoso)"
    }

    if ($script:ClientId -eq "") {
        $script:ClientId = Read-Host "Enter your Application (client) ID (e.g., 7de930ef-77c2-4c63-a7a5-e90d1c2b2521)"
    }

    try {
        $Guid = [Guid]::Parse($script:ClientId)
    }
    catch {
        Write-Host "`nClient ID is invalid or missing. Please enter a valid GUID." -ForegroundColor Red
        Exit
    }
    
    $Url = "https://$($script:TenantName)-admin.sharepoint.com"

    if (($AdminName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential = New-Object System.Management.Automation.PSCredential $AdminName, $SecuredPassword
        Connect-PnPOnline -Url $Url -ClientId $script:ClientId -Credential $Credential
    } elseif ($script:ClientId -ne "" -and $CertificateThumbprint -ne "" -and $script:TenantName -ne "") {
        Connect-PnPOnline -Url $Url -ClientId $script:ClientId -Thumbprint $CertificateThumbprint -Tenant "$script:TenantName.onmicrosoft.com"
    } else {
        Connect-PnPOnline -Url $Url -ClientId $script:ClientId -Interactive
    }
}

# Main function to enable and log external sharing changes
Function Enable-ExternalSharing {

    $Script:EnabledSitesCount = 0
    $Script:RemaingSiteschanged = 0

    Connection-Module

    # Set Tenant-wide sharing settings first
    $TenantSharingCapability = Get-PnPTenant | Select SharingCapability
    
    if($TenantSharingCapability.SharingCapability -ne "ExternalUserAndGuestSharing"){
        Write-Host "You need to enable 'Anyone' sharing at tenant-level to enable external sharing for site(s)."
        $Confirm = Read-Host Are you sure you want to enable 'Anyone' sharing at tenant-level? [Y] Yes [N] No
        if($Confirm -match "[Yy]"){
            Write-Host "`nEnabling 'AnyOne' sharing permission at tenant-level..."
            Set-PnPTenant -SharingCapability "ExternalUserAndGuestSharing" -DefaultSharingLinkType Internal -DefaultLinkPermission View
        }
        else{
            Write-Host "'AnyOne' sharing permission must be enabled at tenant-level to allow external sharing for sites."
            Exit
        }
    }

    # Determine the site collections to be processed from user input
    if ($SiteUrl -ne "") {
        $UserProvidedSites = $SiteUrl
    } elseif ($ImportCsv -ne "") {
        $SiteURLs = Import-Csv -Path $ImportCsv
        $UserProvidedSites = $SiteURLs.SiteUrl
    }
    else{
        $UserProvidedSites = Read-Host "Enter the SharePoint site url"
        if($UserProvidedSites -eq ""){
            Disconnect-PnPOnline
            Exit
        }
    }

    $UserEnteredSites = $UserProvidedSites | Where-Object { $_ -notlike "*-my.sharepoint.com*" }

    $script:LogData = @()

    Function Changes-Summary{
        param([string]$Source)
        if ($Source -eq 'UserEntered') {
            $script:EnabledSitesCount++
        } elseif ($Source -eq 'Remaining') {
            $script:RemaingSiteschanged++
        }
    }

    Function Log-Capability-Changes{

        param([string]$Sites, [string]$SharingMethod, [string]$Source)

             $SharingCapabilityMap = @{
                "Disabled"                    = "Only people in your org"
                "ExternalUserSharingOnly"      = "New and existing guests"
                "ExistingExternalUserSharingOnly" = "Existing guests"
                "ExternalUserAndGuestSharing"  = "Anyone"
             }

            $SiteSharingInfo = Get-PnPTenantSite -Url $Sites | Select-Object SharingCapability, SensitivityLabel, Title
            $SensitivityLabel = $SiteSharingInfo.SensitivityLabel
            $OldSharingCapability = $SiteSharingInfo.SharingCapability.ToString()
            $SiteTitle = $SiteSharingInfo.Title

            $OldSharingCapabilityFriendly = $SharingCapabilityMap[$OldSharingCapability]
            $NewSharingMethodFriendly = $SharingCapabilityMap[$SharingMethod]

            if ($OldSharingCapability -eq $SharingMethod) {
                $script:LogData += "$SiteTitle ($Sites) - External sharing configuration is already set to '$($NewSharingMethodFriendly)'"
                Changes-Summary -Source $Source
            }
            elseif ($SensitivityLabel) {
                $script:LogData += "$SiteTitle ($Sites) - The site inherent a sensitivity label. So, the configuration is not changed."
            }
            else{
                Set-PnPTenantSite -Url $Sites -SharingCapability $SharingMethod
                $script:LogData += "$SiteTitle ($Sites) - External sharing configuration changed from '$($OldSharingCapabilityFriendly)' to '$($NewSharingMethodFriendly)'."
                Changes-Summary -Source $Source
           }
    }

    Write-Host "`nEnabling 'AnyOne' sharing for given site(s)..." -ForegroundColor DarkYellow
    # Process user-entered sites
    foreach ($Site in $UserEnteredSites) {
        Write-Progress -Activity "Processing $Site"
        try {
            Log-Capability-Changes -Sites $Site -SharingMethod "ExternalUserAndGuestSharing" -Source 'UserEntered'
        } catch {
            $script:LogData += "Error processing site '$Site': $_"
        }
    }

    # Process remaining sites: Disable external sharing
    if ($SharingConfigForRemainingSites) {

        $RemainingSites = Get-PnPTenantSite | Select-Object -ExpandProperty Url| Where-Object { $_ -notlike "*-my.sharepoint.com*" -and $UserEnteredSites -notcontains $_ }

        $SharingCapability = switch ($SharingConfigForRemainingSites) {
        'Only people in your org' { "Disabled" }
        'New and existing guests' { "ExternalUserSharingOnly" }
        'Existing guests' { "ExistingExternalUserSharingOnly" }
        'Anyone' { "ExternalUserAndGuestSharing" } 
        }

        Write-Host "`nConfiguring '$SharingConfigForRemainingSites' sharing permission for remaining sites..." -ForegroundColor DarkYellow

        foreach ($Site in $RemainingSites) {
            Write-Progress -Activity "Processing $Site"
            try {
                Log-Capability-Changes -Sites $Site -SharingMethod $SharingCapability -Source 'Remaining'
            } catch {
                $script:LogData += "Error processing site '$Site': $_"
            }
        }
    }

    # Export log if any data is logged
    if ($script:LogData.Count -gt 0) {
        $script:LogData | Out-File -FilePath $LogFilePath -Append
    }

Write-Host "`nSummary of External Sharing Changes" -ForegroundColor DarkYellow
#Write-Host "Sites with sensitivity labels: $script:SensitivityLabeledSitesCount"
Write-Host "External sharing has been enabled on $script:EnabledSitesCount out of $($UserProvidedSites.Count) sites."
if($SharingConfigForRemainingSites){
    Write-Host "$SharingConfigForRemainingSites sharing has been enabled on $script:RemaingSiteschanged out of $($RemainingSites.Count) remaining sites."
}
}

Installation-Module

$script:TenantName = $TenantName
$script:ClientId = $ClientId
$LogFilePath = "$PSScriptRoot\External_Sharing_enabled_site_log$((Get-Date -Format yyyy-MM-dd-ddd` hh-mm-ss` tt).ToString()).txt"

if ($RegisterNewApp.IsPresent) {
    Interactive-App-Register
    Enable-ExternalSharing
} else {
    Enable-ExternalSharing
}
Disconnect-PnPOnline

if ((Test-Path -Path $LogFilePath) -eq "True") {
    Write-Host "`nFor more details, please refer to the log file located at: " -NoNewline -ForegroundColor Yellow
    Write-Host $LogFilePath
    Write-Host "`n~~ Script prepared by AdminDroid Community ~~`n" -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green "`n`n"
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open the output file?", 0, "Open Output File", 4)
    If ($UserInput -eq 6) {
        Invoke-Item "$LogFilePath"
    }
} else {
    Write-Host -ForegroundColor Yellow "No Records Found"
}