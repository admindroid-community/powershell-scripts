
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$Disconnect,
    [ValidateSet('AzureAD','MSOnline','ExchangeOnline','SharePoint','SharePointPnP','SecAndCompCenter','Skype','Teams')]
    [string[]]$Services=("AzureAD","MSOnline","ExchangeOnline",'SharePoint','SharePointPnP','SecAndCompCenter','Skype','Teams'),
    [string]$SharePointHostName,
    [Switch]$MFA,
    [string]$UserName, 
    [string]$Password
)
 
#Disconnecting Sessions
if($Disconnect.IsPresent)
{
 #Disconnect Exchange Online,Skype and Security & Compliance center session
 Get-PSSession | Remove-PSSession
 #Disconnect Teams connection
 Disconnect-MicrosoftTeams
 #Disconnect SharePoint connection
 Disconnect-SPOService
 Write-Host All sessions in the current window has been removed. -ForegroundColor Yellow
}
 
else
{
 if(($UserName -ne "") -and ($Password -ne "")) 
 { 
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force 
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword 
 } 

 #Getting credential for non-MFA account
 elseif(!($MFA.IsPresent)) 
 { 
  $Credential=Get-Credential -Credential $null
 } 
 $ConnectedServices=""
 if($Services.Length -eq 8)
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
  Write-Host Checking connection to $Service...
  Switch ($Service)
  {  
   #Module and Connection settings for Exchange Online module
   ExchangeOnline
   {
    $Module=Get-InstalledModule -Name ExchangeOnlineManagement -MinimumVersion 2.0.3
    if($Module.count -eq 0)
    {
     Write-Host Required Exchange Online'(EXO V2)' module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module ExchangeOnlineManagement
      Import-Module ExchangeOnlineManagement
     }
     else
     {
      Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet.
     }
     Continue
    }
    if($mfa.IsPresent)
    {
     Connect-ExchangeOnline
    }
    else
    {
     Connect-ExchangeOnline -Credential $Credential
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
      Install-Module MSOnline
      Import-Module MSOnline
     }
     else
     {
      Write-Host MSOnline module is required to connect AzureAD.Please install module using Install-Module MSOnline cmdlet.
     }
     Continue
    }
    if($mfa.IsPresent)
    {
     Connect-MsolService
    }
    else
    {
     Connect-MsolService -Credential $Credential
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
      Install-Module AzureAD
      Import-Module AzureAD
     }
     else
     {
      Write-Host AzureAD module is required to connect AzureAD.Please install module using Install-Module AzureAD cmdlet.
     }
     Continue
    }
    if($mfa.IsPresent)
    {
     Connect-AzureAD
    }
    else
    {
     Connect-AzureAD -Credential $Credential
    }
    If((Get-AzureADUser -Top 1) -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+" AzureAD"
     
    }
   }

   #Module and Connection settings for SharePoint Online module
   SharePoint
   {
    $Module=Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable 
    if($Module.count -eq 0)
    {
     Write-Host SharePoint Online PowerShell module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module Microsoft.Online.SharePoint.PowerShell
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
     
    if($MFA.IsPresent)
    {
     Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
     Connect-SPOService -Url https://$SharePointHostName-admin.sharepoint.com
    }
    else
    {
     Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
     Connect-SPOService -Url https://$SharePointHostName-admin.sharepoint.com -credential $credential
    } 
    if((Get-SPOTenant) -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+"SharePoint Online"
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
      Install-Module -Name SharePointPnPPowerShellOnline -AllowClobber
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
     
    if($MFA.IsPresent)
    {
     Import-Module SharepointpnpPowerShellOnline -DisableNameChecking
     Connect-PnPOnline -Url https://$SharePointHostName.sharepoint.com -UseWebLogin -WarningAction Ignore
    }
    else
    {
     Import-Module SharepointpnpPowerShellOnline -DisableNameChecking
     Connect-PnPOnline -Url https://$SharePointHostName.sharepoint.com -credential $credential -WarningAction Ignore
    } 
    If ($? -eq $true)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+"SharePoint PnP"  
    }
   }


   #Module and Connection settings for Skype for Business Online module
   Skype
   { 
    $Module=Get-InstalledModule -Name MicrosoftTeams -MinimumVersion 1.1.6 
    if($Module.count -eq 0)
    {
     Write-Host Required MicrosoftTeams module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module MicrosoftTeams -AllowClobber
     }
     else
     {
      Write-Host MicrosoftTeams module is required.Please install module using Install-Module MicrosoftTeams cmdlet.
     }
     Continue
    }
    if($MFA.IsPresent)
    {
     $sfbSession = New-CsOnlineSession
     Import-PSSession $sfbSession -AllowClobber | Out-Null
    }
    else
    {
     $sfbSession = New-CsOnlineSession -Credential $Credential
     Import-PSSession $sfbSession -AllowClobber -WarningAction SilentlyContinue | Out-Null
    }
    #Check for Skype connectivity
    If ((Get-PSSession | Where-Object { $_.ConfigurationName -like "Microsoft.PowerShell" }) -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+"Skype"  
    }
   }

   #Module and Connection settings for Security & Compliance center
   SecAndCompCenter
   {
    $Module=Get-InstalledModule -Name ExchangeOnlineManagement -MinimumVersion 2.0.3
    if($Module.count -eq 0)
    {
     Write-Host Exchange Online'(EXO V2)' module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module ExchangeOnlineManagement
      Import-Module ExchangeOnlineManagement
     }
     else
     {
      Write-Host EXO V2 module is required to connect Security and Compliance PowerShell.Please install module using Install-Module ExchangeOnlineManagement cmdlet.
     }
     Continue
    }
    if($mfa.IsPresent)
    {
     Connect-IPPSSession -WarningAction SilentlyContinue
    }
    else
    {
     Connect-IPPSSession -Credential $Credential -WarningAction SilentlyContinue
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
   Teams
   {
    $Module=Get-InstalledModule -Name MicrosoftTeams -MinimumVersion 1.1.6 
    if($Module.count -eq 0)
    {
     Write-Host Required MicrosoftTeams module is not available  -ForegroundColor yellow 
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
     if($Confirm -match "[yY]")
     {
      Install-Module MicrosoftTeams -AllowClobber
     }
     else
     {
      Write-Host MicrosoftTeams module is required.Please install module using Install-Module MicrosoftTeams cmdlet.
     }
     Continue
    }
    if($mfa.IsPresent)
    {
     $Team=Connect-MicrosoftTeams
    }
    else
    {
     $Team=Connect-MicrosoftTeams -Credential $Credential
    }
    #Check for Teams connectivity
    If($Team -ne $null)
    {
     if($ConnectedServices -ne "")
     {
      $ConnectedServices=$ConnectedServices+","
     }
     $ConnectedServices=$ConnectedServices+"Teams"
    }
   }
  }
 }
 if($ConnectedServices -eq "")
 {
  $ConnectedServices="-"
 }
 Write-Host `n`nConnected Services $ConnectedServices -ForegroundColor DarkYellow 
}
