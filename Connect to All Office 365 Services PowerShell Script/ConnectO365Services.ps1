<#
=============================================================================================
Name:           Connect to all the Microsoft services using PowerShell
Description:    This script automatically installs all the required modules(upon your confirmation) and connects to the services
Version:        5.0
Website:        o365reports.com

Script Highlights:
~~~~~~~~~~~~~~~~~

1.This script connects to 9 Microsoft 365 services with a single cmdlet.
2.Installs Microsoft 365 PowerShell modules. ie, Modules required for Microsoft 365 services are automatically downloaded and installed upon your confirmation.
3.You can connect to one or more Microsoft 365 services via PowerShell using a single cmdlet.
4.You can connect to Microsoft 365 services with MFA enabled account.
5.For non-MFA account, you don’t need to enter credential for each service.
5.The script is scheduler friendly. i.e., credentials can be passed as a parameter instead of saving inside the script.
6.You can disconnect all service connections using a single cmdlet.
7.The script supports Certificate-Based Authentication (CBA) too.

For detailed script execution: https://o365reports.com/2019/10/05/connect-all-office-365-services-powershell/

Change Log:
~~~~~~~~~~~
~~~~~~~~~
  V1.0 (Nov 01, 2019) - File created
  V2.0 (Jan 21, 2020) - Added support for MS Online and SharePoint PnP PowerShell modules
  V3.0 (Oct 06, 2023) - Removed Skype for Business and minor usability changes
  V4.0 (Feb 29, 2024) - Added support for MS Graph and MS Graph beta PowerShell modules
  V4.1 (Apr 03, 2025) - Handled ClientId requirement for SharePoint PnP PowerShell module
  V5.0 (Dec 26, 2025) - Removed MSOnline & AzureAD modules and added MS Entra module.
                        Included CBA for SPOService module and made minor usability changes

============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$Disconnect,
    # Load order needs MSGraph early to obtain SharePointHostName. MSTeams also needs to be early in the load order to prevent login issues
    [ValidateSet('MSGraph','MSGraphBeta','MSTeams','SharePointOnline','SharePointPnP','SecAndCompCenter','ExchangeOnline','MSEntra')]
    [string[]]$Services=('MSGraph','MSGraphBeta','MSTeams','SharePointOnline','SharePointPnP','SecAndCompCenter','ExchangeOnline','MSEntra'),
    [string]$SharePointHostName,
    [Switch]$MFA,
    [Switch]$CBA,
    [string]$TenantId,
    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$UserName,
    [string]$Password,
    [string[]]$GraphScopes=("User.Read.All"),
    [string[]]$EntraScopes=("User.Read.All")
)

#Disconnecting Sessions
if($Disconnect.IsPresent)
{
 #Disconnect Exchange Online and Security & Compliance center connection
 $dummy = Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
 #Disconnect Teams connection
 $dummy = Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
 #Disconnect SharePoint/PnP connection
 $dummy = Disconnect-SPOService -ErrorAction SilentlyContinue
 $dummy = Disconnect-PnPOnline -ErrorAction SilentlyContinue
 #Disconnect MS Graph PowerShell connection
 $dummy = Disconnect-MgGraph -ErrorAction SilentlyContinue
 #Disconnect MS Entra PowerShell connection
 $dummy = Disconnect-Entra -ErrorAction SilentlyContinue

 Write-Host All sessions in the current window have been removed. -ForegroundColor Yellow
}
else
{
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  $CredentialPassed=$true
 }
 elseif(($AppId -ne "") -and ($CertificateThumbPrint -ne "") -and ($TenantId -ne ""))
 {
  $CBA=$true
 }
 else
 {
  $MFA=$true
 }

 if($CBA -eq $true)
 {
  $Certificate = Get-ChildItem Cert:\ -Recurse |
   Where-Object {
    $_.Thumbprint -eq $CertificateThumbprint -and
    $_.HasPrivateKey
   } |
   Select-Object -First 1
  if (-not $Certificate)
  {
   Write-Host Certificate with thumbprint $CertificateThumbprint not found in certificate store -ForegroundColor Red
   Write-Host Falling back to MFA. -ForegroundColor Yellow
   $CBA=$false
   $MFA=$true
  }
  else
  {
   Write-Host Retrieved certificate from store -ForegroundColor Green
  }
 }

 $ConnectedServices=@()
 $RequiredServices=$Services

 $InstalledModules = @{}
 Get-InstalledModule -ErrorAction SilentlyContinue | ForEach-Object {
  $InstalledModules[$_.Name] = $_
 }

 #Loop through each required services
 Foreach($Service in $RequiredServices)
 {
  Write-Host Connecting to $Service... -ForegroundColor Green
  Switch ($Service)
  {
   #Module and Connection settings for Exchange Online module
   ExchangeOnline
   {
    if(-not $InstalledModules.ContainsKey('ExchangeOnlineManagement'))
    {
     Write-Host Required Exchange Online PowerShell module is not available  -ForegroundColor yellow
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -eq 'y')
     {
      Install-Module ExchangeOnlineManagement -Scope CurrentUser
     }
     else
     {
      Write-Host EXO PowerShell module is required to connect Exchange Online. Please install module using Install-Module ExchangeOnlineManagement cmdlet.
     }
     Continue
    }

    if($CredentialPassed -eq $true)
    {
     Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
    }
    elseif($CBA -eq $true)
    {
     if(!($PSBoundParameters['SharePointHostName']) -and ([string]$SharePointHostName -eq "") )
     {
      Write-Host SharePoint organization name is required.`nEg: Contoso for admin@Contoso.Onmicrosoft.com -ForegroundColor Yellow
      $SharePointHostName = Read-Host "Please enter SharePoint organization name"
     }
     Connect-ExchangeOnline -Organization "$($SharePointHostName).onmicrosoft.com" -AppId $AppId -Certificate $Certificate -ShowBanner:$false
    }
    elseif($MFA -eq $true)
    {
     Connect-ExchangeOnline -ShowBanner:$false
    }
    if(Get-ConnectionInformation | Where-Object {$_.State -eq 'Connected' -and -not $_.IsEopSession})
    {
     $ConnectedServices+="Exchange Online"
    }
   }

   #Module and Connection settings for SharePoint Online module
   SharePointOnline
   {
    if(-not $InstalledModules.ContainsKey('Microsoft.Online.SharePoint.PowerShell'))
    {
     Write-Host SharePoint Online PowerShell module is not available  -ForegroundColor yellow
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -eq 'y')
     {
      Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
     }
     else
     {
      Write-Host SharePoint Online PowerShell module is required. Please install module using Install-Module Microsoft.Online.SharePoint.PowerShell cmdlet.
     }
     Continue
    }

    if(!($PSBoundParameters['SharePointHostName']) -and ([string]$SharePointHostName -eq "") )
    {
     Write-Host SharePoint organization name is required.`nEg: Contoso for admin@Contoso.Onmicrosoft.com -ForegroundColor Yellow
     $SharePointHostName = Read-Host "Please enter SharePoint organization name"
    }

    if(($PSVersionTable::PSVersion.Major) -ge 7)
    {
     Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue
     Write-Host The login dialog could be hidden behind another window -ForegroundColor Red
    }

    if($CredentialPassed -eq $true)
    {
     Connect-SPOService -Url https://$($SharePointHostName)-admin.sharepoint.com -credential $credential
    }
    elseif($CBA -eq $true)
    {
     # This module does not support Certificate auth in -UseWindowsPowerShell mode in Powershell 7. Falling back to MFA mode
     if(($PSVersionTable::PSVersion.Major) -ge 7)
     {
      Connect-SPOService -Url https://$($SharePointHostName)-admin.sharepoint.com
     }
     else
     {
      Connect-SPOService -Url https://$($SharePointHostName)-admin.sharepoint.com -TenantId $TenantId -ClientId $AppId -Certificate $Certificate
     }
    }
    elseif($MFA -eq $true)
    {
     Connect-SPOService -Url https://$($SharePointHostName)-admin.sharepoint.com
    }
    if((Get-SPOTenant) -ne $null)
    {
     $ConnectedServices+="SharePoint Online"
    }
   }

   #Module and Connection settings for Sharepoint PnP module
   SharePointPnP
   {
    # PnP.PowerShell 1.12.0 was the last to support Powershell 5.
    if(($PSVersionTable::PSVersion.Major) -ge 7)
    {
     $PnPModule = "PnP.PowerShell" # powershell 7.
    }
    else
    {
     $PnPModule = "SharePointPnPPowerShellOnline" # powershell 5
    }
    if(-not $InstalledModules.ContainsKey($PnPModule))
    {
     Write-Host SharePoint PnP module module is not available  -ForegroundColor yellow
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -eq 'y')
     {
      Install-Module -Name $PnPModule -AllowClobber -Scope CurrentUser
     }
     else
     {
      Write-Host SharePoint PnP module is required. Please install module using Install-Module $PnPModule cmdlet.
     }
     Continue
    }

    if(!($PSBoundParameters['SharePointHostName']) -and ([string]$SharePointHostName -eq "") )
    {
     Write-Host SharePoint organization name is required.`nEg: Contoso for admin@Contoso.Onmicrosoft.com -ForegroundColor Yellow
     $SharePointHostName = Read-Host "Please enter SharePoint organization name"
    }

    if($CBA -eq $true)
    {
     # Powershell 7 PnP module does not support -Certificate. So required to use -Thumbprint.
     Connect-PnPOnline -Url https://$($SharePointHostName)-admin.sharepoint.com -Tenant $TenantId -ClientId $AppId -Thumbprint $CertificateThumbprint -WarningAction Ignore
    }
    elseif(($CredentialPassed -eq $true) -or ($MFA -eq $true))
    {
     if(($PSVersionTable::PSVersion.Major) -ge 7)
     {
      $JSON = Get-Content "$PSScriptRoot/Tenants.json" -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json

      if (-not $JSON) {
       $JSON = [PSCustomObject]@{}
      }

      if (-not $JSON.PnP7) {
       $JSON | Add-Member -MemberType NoteProperty -Name PnP7 -Value ([PSCustomObject]@{})
       #$JSON.PnP7 = @()
      }

      if($JSON.PnP7.$SharePointHostName) {
       $PnPClientID = $JSON.PnP7.$SharePointHostName
      }
      else {
       if($ConnectedServices -contains "MS Graph") {
        $PnPapp = Get-MgApplication -ConsistencyLevel eventual -Filter "DisplayName eq 'PnP.PowerShell'"
       }
       else {
        $PnPapp = $false
       }
       if (-not $PnPapp) {
        Write-Host "Adding PnP.PowerShell App..." -ForegroundColor Yellow
        $NEWPnPapp = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant "$($SharePointHostName).onmicrosoft.com"
        $PnPClientID= $NEWPnPapp.'AzureAppId/ClientId'
       }
       else {
        Write-Host "Existing PnP.PowerShell App found..." -ForegroundColor Yellow
        $PnPClientID = $PnPapp.AppId
       }
       $JSON.PnP7 | Add-Member -MemberType NoteProperty -Name $SharePointHostName -Value $PnPClientID
       #$JSON.PnP7.$SharePointHostName = $PnPClientID
      }

      $JSON | ConvertTo-Json | Set-Content -Path "$PSScriptRoot/Tenants.json"
     }

     if($CredentialPassed -eq $true)
     {
      Connect-PnPOnline -Url https://$($SharePointHostName)-admin.sharepoint.com  -credential $credential -ClientId $PnPClientID  -WarningAction Ignore
     }
     elseif($MFA -eq $true)
     {
      if(($PSVersionTable::PSVersion.Major) -ge 7)
      {
       Connect-PnPOnline -Url https://$($SharePointHostName)-admin.sharepoint.com -Interactive -ClientId $PnPClientID -WarningAction Ignore
      }
      else
      {
       Connect-PnPOnline -Url https://$($SharePointHostName)-admin.sharepoint.com -UseWebLogin -WarningAction Ignore
      }
     }
    }

    if (Get-PnPConnection)
    {
     $ConnectedServices+="SharePoint PnP"
    }
   }

   #Module and Connection settings for Security & Compliance center
   SecAndCompCenter
   {
    if(-not $InstalledModules.ContainsKey('ExchangeOnlineManagement'))
    {
     Write-Host Exchange Online PowerShell module is not available  -ForegroundColor yellow
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -eq 'y')
     {
      Install-Module ExchangeOnlineManagement -Scope CurrentUser
     }
     else
     {
      Write-Host EXO PowerShell module is required to connect Security and Compliance PowerShell. Please install module using Install-Module ExchangeOnlineManagement cmdlet.
     }
     Continue
    }

    if($CredentialPassed -eq $true)
    {
     Connect-IPPSSession -Credential $Credential -ShowBanner:$false
    }
    elseif($CBA -eq $true)
    {
     if(!($PSBoundParameters['SharePointHostName']) -and ([string]$SharePointHostName -eq "") )
     {
      Write-Host SharePoint organization name is required.`nEg: Contoso for admin@Contoso.Onmicrosoft.com -ForegroundColor Yellow
      $SharePointHostName = Read-Host "Please enter SharePoint organization name"
     }
     Connect-IPPSSession -Organization "$($SharePointHostName).onmicrosoft.com" -AppId $AppId -Certificate $Certificate -ShowBanner:$false
    }
    elseif($MFA -eq $true)
    {
     Connect-IPPSSession -ShowBanner:$false
    }
    if(Get-ConnectionInformation | Where-Object {$_.State -eq 'Connected' -and $_.IsEopSession})
    {
     $ConnectedServices+="Security & Compliance Center"
    }
   }

   #Module and Connection settings for Teams Online module
   MSTeams
   {
    if(-not $InstalledModules.ContainsKey('MicrosoftTeams'))
    {
     Write-Host Required MicrosoftTeams module is not available  -ForegroundColor yellow
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -eq 'y')
     {
      Install-Module MicrosoftTeams -AllowClobber -Scope CurrentUser
     }
     else
     {
      Write-Host MicrosoftTeams module is required. Please install module using Install-Module MicrosoftTeams cmdlet.
     }
     Continue
    }

    if($CredentialPassed -eq $true)
    {
     $Teams=Connect-MicrosoftTeams -Credential $Credential
    }
    elseif($CBA -eq $true)
    {
     $Teams=Connect-MicrosoftTeams -TenantId $TenantId -ApplicationId $AppId -Certificate $Certificate
    }
    elseif($MFA -eq $true)
    {
     $Teams=Connect-MicrosoftTeams
    }

    #Check for Teams connectivity
    if(Get-CsTenant)
    {
     $ConnectedServices+="MS Teams"
    }
   }

   #Module and connection settings for MS Graph PowerShell
   MSGraph
   {
    #Check for module installation
    if(-not $InstalledModules.ContainsKey('Microsoft.Graph'))
    {
     Write-Host Microsoft Graph PowerShell SDK is not available  -ForegroundColor yellow
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -eq 'y')
     {
      Write-host "Installing Microsoft Graph PowerShell module..."
      Install-Module Microsoft.Graph -AllowClobber -Scope CurrentUser
     }
     else
     {
      Write-Host "Microsoft Graph PowerShell module is required. Please install module using Install-Module Microsoft.Graph cmdlet."
     }
     Continue
    }

    if($CredentialPassed -eq $true)
    {
     Write-Host "MS Graph doesn't support passing credential as parameters. Please enter the credential in the prompt."
     Connect-MgGraph -Scopes $GraphScopes -ContextScope Process -NoWelcome
    }
    elseif($CBA -eq $true)
    {
     Connect-MgGraph -TenantId $TenantId -ClientId $AppId -Certificate $Certificate -ContextScope Process -NoWelcome
    }
    elseif($MFA -eq $true)
    {
     Connect-MgGraph -Scopes $GraphScopes -ContextScope Process -NoWelcome
    }

    #Check for MS Graph connectivity
    if(Get-MgContext)
    {
     $ConnectedServices+="MS Graph"

     if(!($PSBoundParameters['SharePointHostName']) -and ([string]$SharePointHostName -eq "") )
     {
      $domain = (Get-MgDomain | Where-Object {$_.IsInitial -eq $true}).Id -split ".onmicrosoft.com"
      $SharePointHostName = $domain[0].trim()
     }
    }
   }

   #Module and connection settings for MS Graph Beta PowerShell
   MSGraphBeta
   {
    #Check for module installation
    if(-not $InstalledModules.ContainsKey('Microsoft.Graph.Beta'))
    {
     Write-Host Microsoft Graph Beta PowerShell SDK is not available  -ForegroundColor yellow
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -eq 'y')
     {
      Write-host "Installing Microsoft Graph Beta PowerShell module..."
      Install-Module Microsoft.Graph.Beta -AllowClobber -Scope CurrentUser
     }
     else
     {
      Write-Host "Microsoft Graph Beta PowerShell module is required. Please install module using Install-Module Microsoft.Graph.Beta cmdlet."
     }
     Continue
    }

    # We do not need to login if MS Graph is logged in
    if($ConnectedServices -contains "MS Graph")
    {
     Write-Host "Skipping MS Graph Beta login since MS Graph is already logged in" -ForegroundColor yellow
    }
    else
    {
     if($CredentialPassed -eq $true)
     {
      Write-Host "MS Graph Beta doesn't support passing credential as parameters. Please enter the credential in the prompt."
      Connect-MgGraph -Scopes $GraphScopes -ContextScope Process -NoWelcome
     }
     elseif($CBA -eq $true)
     {
      Connect-MgGraph -TenantId $TenantId -ClientId $AppId -Certificate $Certificate -ContextScope Process -NoWelcome
     }
     elseif($MFA -eq $true)
     {
      Connect-MgGraph -Scopes $GraphScopes -ContextScope Process -NoWelcome
     }
    }

    #Check for MS Graph Beta connectivity
    if(Get-MgContext)
    {
     $ConnectedServices+="MS Graph Beta"

     if(!($PSBoundParameters['SharePointHostName']) -and ([string]$SharePointHostName -eq "") )
     {
      $domain = (Get-MgBetaDomain | Where-Object {$_.IsInitial -eq $true}).Id -split ".onmicrosoft.com"
      $SharePointHostName = $domain[0].trim()
     }
    }
   }

   #Module and connection settings for MS Entra PowerShell
   MSEntra
   {
    #Check for module installation
    if(-not $InstalledModules.ContainsKey('Microsoft.Entra'))
    {
     Write-Host Microsoft Entra PowerShell module is not available  -ForegroundColor yellow
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -eq 'y')
     {
      Write-host "Installing Microsoft Entra PowerShell module..."
      Install-Module Microsoft.Entra -AllowClobber -Scope CurrentUser
     }
     else
     {
      Write-Host "Microsoft Entra PowerShell module is required. Please install module using 'Install-Module Microsoft.Entra' cmdlet."
     }
     Continue
    }

    # We do not need to login if MS Graph is logged in
    if($ConnectedServices -contains "MS Graph")
    {
     Write-Host "Skipping MS Entra login since MS Graph is already logged in" -ForegroundColor yellow
    }
    else
    {
     if($CredentialPassed -eq $true)
     {
      Write-Host "MS Entra doesn't support passing credential as parameters. Please enter the credential in the prompt."
      Connect-Entra -Scopes $EntraScopes -ContextScope Process -NoWelcome
     }
     elseif($CBA -eq $true)
     {
      Connect-Entra -ApplicationId $AppId -TenantId $TenantId -Certificate $Certificate -ContextScope Process -NoWelcome
     }
     elseif($MFA -eq $true)
     {
      Connect-Entra -Scopes $EntraScopes -ContextScope Process -NoWelcome
     }
    }

    #Check for MS Entra connectivity
    if((Get-EntraUser -Top 1) -ne $null) # Get-EntraContext for entra 1.3 or higher
    {
     $ConnectedServices+="MS Entra"
    }
   }
  }
 }

 $CSstring = $ConnectedServices -join ", "
 Write-Host `n`nConnected Services - $CSstring -ForegroundColor Cyan
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green `n`n
}

