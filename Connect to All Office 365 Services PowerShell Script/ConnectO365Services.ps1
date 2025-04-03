<#
=============================================================================================
Name:           Connect to all the Microsoft services using PowerShell
Description:    This script automatically installs all the required modules(upon your confirmation) and connects to the services
Version:        4.1
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
  V2.0 (Jan 21, 2020)  - Added support for MS Online and SharePoint PnP PowerShell modules
  V3.0 (Oct 06, 2023)  - Removed Skype for Business and minor usability changes
  V4.0 (Feb 29, 2024) - Added support for MS Graph and MS Graph beta PowerShell modules
  V4.1 (Apr 03, 2025) - Handled ClientId requirement for SharePoint PnP PowerShell module
  
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$Disconnect,
    [ValidateSet('MSGraph','MSGraphBeta','ExchangeOnline','SharePointOnline','SharePointPnP','SecAndCompCenter','MSTeams','MSOnline','AzureAD')]
    [string[]]$Services=("ExchangeOnline",'MSTeams','SharePointOnline','SharePointPnP','SecAndCompCenter',"MSOnline","AzureAD",'MSGraph','MSGraphBeta'),
    [string]$SharePointHostName,
    [Switch]$MFA,
    [Switch]$CBA,
    [String]$TenantName,
    [string]$TenantId,
    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$UserName, 
    [string]$Password
)
 
#Disconnecting Sessions
if($Disconnect.IsPresent)
{
 #Disconnect Exchange Online,Skype and Security & Compliance center session
 Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
 #Disconnect Teams connection
 Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
 #Disconnect SharePoint connection
 Disconnect-SPOService -ErrorAction SilentlyContinue
 Disconnect-PnPOnline -ErrorAction SilentlyContinue
 #Disconnect MS Graph PowerShell
 Disconnect-MgGraph -ErrorAction SilentlyContinue
 Write-Host All sessions in the current window has been removed. -ForegroundColor Yellow
}
 
else
{
 if(($UserName -ne "") -and ($Password -ne "")) 
 { 
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force 
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword 
  $CredentialPassed=$true
 } 
 elseif((($AppId -ne "") -and ($CertificateThumbPrint -ne "")) -and (($TenantId -ne "") -or ($TenantName -ne "")))
 {
  $CBA=$true
 }

 $ConnectedServices=""
 if($Services.Length -eq 9)
 {
  $RequiredServices=$Services  
 }
 else
 {
  $RequiredServices=$PSBoundParameters.Services
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
    $Module=Get-InstalledModule -Name ExchangeOnlineManagement
    if($Module.count -eq 0)
    {
     Write-Host Required Exchange Online PowerShell module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module ExchangeOnlineManagement -Scope CurrentUser
      Import-Module ExchangeOnlineManagement
     }
     else
     {
      Write-Host EXO PowerShell module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet.
     }
     Continue
    }
  
    if($CredentialPassed -eq $true)
    {
     Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
    }
    elseif($CBA -eq $true)
    {
     Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint  -Organization $TenantName -ShowBanner:$false
    }
    else
    {
     Connect-ExchangeOnline -ShowBanner:$false
    }
    If((Get-EXOMailbox -ResultSize 1) -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+" Exchange Online"
    }
   }

  #Module and Connection settings for AzureAD V1 (MSOnline module)
   MSOnline
   {
    $Module=Get-Module -Name MSOnline -ListAvailable 
    if($Module.count -eq 0)
    {
     Write-Host MSOnline module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module MSOnline -Scope CurrentUser
      Import-Module MSOnline
     }
     else
     {
      Write-Host MSOnline module is required to connect AzureAD.Please install module using Install-Module MSOnline cmdlet.
     }
     Continue
    }

    if($CredentialPassed -eq $true)
    {
     Connect-MsolService -Credential $Credential
    }
    elseif($CBA -eq $true)
    {
     Write-Host "MSonline module doesn't support certificate based authentication. Please enter the credential in the prompt"
     Connect-MsolService
    }
    else
    {
     Connect-MsolService
    }
    If((Get-MsolUser -MaxResults 1) -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+" MSOnline"
    }
    if(($RequiredServices -contains "SharePoint") -eq "true")
     {
      $SharePointHostName=((Get-MsolDomain) | where {$_.IsInitial -eq "True"} ).name -split ".onmicrosoft.com"
      $SharePointHostName=($SharePointHostName).trim()
     }
   }



   #Module and Connection settings for AzureAD V2 (AzureAD module)
   AzureAD
   {
    $Module=Get-Module -Name AzureAD -ListAvailable 
    if($Module.count -eq 0)
    {
     Write-Host AzureAD module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module AzureAD -Scope CurrentUser
      Import-Module AzureAD
     }
     else
     {
      Write-Host AzureAD module is required to connect AzureAD.Please install module using Install-Module AzureAD cmdlet.
     }
     Continue
    }
    
    if($CredentialPassed -eq $true)
    {
     $AzureAD=Connect-AzureAD -Credential $Credential
    }
    elseif($CBA -eq $true)
    {
     $AzureAD=Connect-AzureAD -ApplicationId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
    }
    else
    {
     $AzureAD=Connect-AzureAD
    }

    #Check for Azure AD connectivity
    If($AzureAD -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+" Azure AD"
    }
   }

   #Module and Connection settings for SharePoint Online module
   SharePointOnline
   {
    $Module=Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable 
    if($Module.count -eq 0)
    {
     Write-Host SharePoint Online PowerShell module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
      Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
     }
     else
     {
      Write-Host SharePoint Online PowerShell module is required.Please install module using Install-Module Microsoft.Online.SharePoint.PowerShell cmdlet.
      Continue
     }
    }
    if(!($PSBoundParameters['SharePointHostName']) -and ([string]$SharePointHostName -eq "") ) 
    {
     Write-Host SharePoint organization name is required.`nEg: Contoso for admin@Contoso.Onmicrosoft.com -ForegroundColor Yellow
     $SharePointHostName= Read-Host "Please enter SharePoint organization name"  
    }

    if($CredentialPassed -eq $true)
    {
     Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
     Connect-SPOService -Url https://$SharePointHostName-admin.sharepoint.com -credential $credential
    } 
    elseif($CBA -eq $true)
    {
     Write-Host "SharePoint Online PowerShell module doesn't support certificate based authentication. Please enter credential in the prompt"
     Connect-SPOService -Url https://$SharePointHostName-admin.sharepoint.com
    }
    else
    {
     Connect-SPOService -Url https://$SharePointHostName-admin.sharepoint.com
    }
    if((Get-SPOTenant) -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+" SharePoint Online"
    }
   }

   #Module and Connection settings for Sharepoint PnP module
   SharePointPnP
   {
    $Module=Get-InstalledModule -Name SharePointPnPPowerShellOnline
    if($Module.count -eq 0)
    {
     Write-Host SharePoint PnP module module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module -Name SharePointPnPPowerShellOnline -AllowClobber -Scope CurrentUser
      Import-Module SharepointpnpPowerShellOnline -DisableNameChecking
     }
     else
     {
      Write-Host SharePoint Pnp module is required.Please install module using Install-Module SharePointPnPPowerShellOnline cmdlet.
     }
     Continue
    }
    if(!($PSBoundParameters['SharePointHostName']) -and ([string]$SharePointHostName -eq "") ) 
    {
     Write-Host SharePoint organization name is required.`nEg: Contoso for admin@Contoso.com -ForegroundColor Yellow
     $SharePointHostName= Read-Host "Please enter SharePoint organization name"  
    }
    
    if($AppId -eq "")
    {
     Write-Host `nClient Id is mandatory to connect SharePoint PnP PowerShell -ForegroundColor Yellow
     $AppId= Read-Host "Please enter client id to connect PnP PowerShell"
    }
   
    if($CredentialPassed -eq $true)
    {
     Connect-PnPOnline -Url https://$SharePointHostName-admin.sharepoint.com  -credential $credential -ClientId $AppId  -WarningAction Ignore
    } 
    elseif($CBA -eq $true)
    {
     if($TenantName -eq "")
     {
      Write-Host Tenant name is required.`ne.g. contoso.onmicrosoft.com -ForegroundColor Yellow
      $TenantName= Read-Host "Please enter your tenant name"  
     }
     Connect-PnPOnline -Url https://$SharePointHostName-admin.sharepoint.com -ClientId $AppId -Thumbprint $CertificateThumbprint -Tenant $TenantName
    }
     else
    {
     
     Connect-PnPOnline -Url https://$SharePointHostName-admin.sharepoint.com -ClientId $AppId -WarningAction Ignore -Interactive
    }
    If ($? -eq $true)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+" SharePoint PnP"  
    }
   }

   #Module and Connection settings for Security & Compliance center
   SecAndCompCenter
   {
    $Module=Get-InstalledModule -Name ExchangeOnlineManagement
    if($Module.count -eq 0)
    {
     Write-Host Exchange Online PowerShell module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module ExchangeOnlineManagement -Scope CurrentUser
      Import-Module ExchangeOnlineManagement
     }
     else
     {
      Write-Host EXO PowerShell module is required to connect Security and Compliance PowerShell.Please install module using Install-Module ExchangeOnlineManagement cmdlet.
     }
     Continue
    }
 
    if($CredentialPassed -eq $true)
    {
     Connect-IPPSSession -Credential $Credential -ShowBanner:$false
    }
    elseif($CBA -eq $true)
    {
     if($TenantName -eq "")
     {
      Write-Host Orgnaization name is required.`ne.g. contoso.onmicrosoft.com -ForegroundColor Yellow
      $TenantName= Read-Host "Please enter your Organization name"  
     }
     Connect-IPPSSession -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $TenantName -ShowBanner:$false
    }
    else
    {
     Connect-IPPSSession -ShowBanner:$false
    }
    $Result=Get-RetentionCompliancePolicy
    If(($?) -eq $true)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+" Security & Compliance Center"
    }
   }
  
   #Module and Connection settings for Teams Online module
  MSTeams
   {
    $Module=Get-InstalledModule -Name MicrosoftTeams -MinimumVersion 4.0.0 
    if($Module.count -eq 0)
    {
     Write-Host Required MicrosoftTeams module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module MicrosoftTeams -AllowClobber -Force -Scope CurrentUser
      Import-Module MicrosoftTeams
     }
     else
     {
      Write-Host MicrosoftTeams module is required.Please install module using Install-Module MicrosoftTeams cmdlet.
     }
     Continue
    }

    if($CredentialPassed -eq $true)
    {
     $Teams=Connect-MicrosoftTeams -Credential $Credential
    }
    elseif($CBA -eq $true)
    {
     $Teams=Connect-MicrosoftTeams -ApplicationId $AppId -TenantId $TenantId -CertificateThumbPrint $CertificateThumbprint
    }
    else
    {
     $Teams=Connect-MicrosoftTeams
    }

    #Check for Teams connectivity
    If($Teams -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+" MS Teams"
    }
   }

   #Module and connection settings for MS Graph PowerShell
  MSGraph
   {
    #Check for module installation
    $Module=Get-Module -Name microsoft.graph -ListAvailable
    if($Module.count -eq 0) 
    { 
     Write-Host Microsoft Graph PowerShell SDK is not available  -ForegroundColor yellow  
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
     if($Confirm -match "[yY]") 
     { 
      Write-host "Installing Microsoft Graph PowerShell module..."
      Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
      Import-Module Microsoft.Graph.Users
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
     Connect-MgGraph -NoWelcome
    }
    elseif($CBA -eq $true)
    {
     Connect-MgGraph -ApplicationId $AppId -TenantId $TenantId -CertificateThumbPrint $CertificateThumbprint -NoWelcome
    }
    else
    {
     Connect-MgGraph -NoWelcome
    }

    #Check for MS Graph connectivity
 If((Get-MgUser -Top 1) -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+" MS Graph"
     
    }
   }

   #Module and connection settings for MS Graph Beta PowerShell
  MSGraphBeta
   {
    #Check for module installation
    $Module=Get-Module -Name microsoft.graph.beta -ListAvailable
    if($Module.count -eq 0) 
    { 
     Write-Host Microsoft Graph Beta PowerShell SDK is not available  -ForegroundColor yellow  
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
     if($Confirm -match "[yY]") 
     { 
      Write-host "Installing Microsoft Graph Beta PowerShell module..."
      Install-Module Microsoft.Graph.Beta -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
      Import-Module Microsoft.Graph.Beta.Users
     }
     else
     {
      Write-Host "Microsoft Graph Beta PowerShell module is required. Please install module using Install-Module Microsoft.Graph.Beta cmdlet." 
     }
     Continue
    }
   
    if($CredentialPassed -eq $true)
    {
     Write-Host "MS Graph Beta doesn't support passing credential as parameters. Please enter the credential in the prompt."
     Connect-MgGraph -NoWelcome
    }
    elseif($CBA -eq $true)
    {
     Connect-MgGraph -ApplicationId $AppId -TenantId $TenantId -CertificateThumbPrint $CertificateThumbprint -NoWelcome
    }
    else
    {
     Connect-MgGraph -NoWelcome
    }

    #Check for MS Graph Beta connectivity
    If((Get-MgBetaUser -Top 1) -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+" MS Graph Beta"
     
    }
   }
  }
 }
 if($ConnectedServices -eq "")
 {
  $ConnectedServices="-"
 }
 Write-Host `n`nConnected Services - $ConnectedServices -ForegroundColor Cyan
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
}

