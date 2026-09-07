<#
=============================================================================================
Name:           Get App registrations and their owner, client secrets, & certificates
Description:    The script exports all app regsitrations and their client secrets & certificates along with the expiry status, Expiry date, etc.
Version:        1.0
Website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the Microsoft Graph PowerShell SDK module (if not installed already) upon your confirmation. 
2. Exports certificates and secrets for all app registrations in Microsoft Entra. 
3. Generates a report that retrieves certificates from apps. 
4. Exports a report that retrieves client secrets from apps. 
5. Identifies applications with only active certificates or client secrets. 
6. Lists applications that have only expired certificates or client secrets. 
7. Allows to export the list of apps soon to expire client secrets and certificates (i.e., 30 days, 90 days, etc.) 
8. Exports the report result to CSV. 
9. The script can be executed with an MFA enabled account too. 
10. It can be executed with Certificate-based Authentication (CBA) too. 
11. The script is schedular-friendly.

For detailed Script execution: 

https://o365reports.com/2025/02/25/export-all-app-registrations-with-certificates-and-secrets-in-microsoft-entra

============================================================================================
#>


Param
(
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [Switch]$ClientSecretsOnly,
    [Switch]$CertificatesOnly,
    [Switch]$ExpiredOnly,
    [Switch]$ActiveOnly,
    [int]$SoonToExpireInDays
)
Function Connect_MgGraph
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
  }
  else
  {
   Write-Host "Microsoft Graph PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph cmdlet." 
   Exit
  }
 }
 #Disconnect Existing MgGraph session
 if($CreateSession.IsPresent)
 {
  Disconnect-MgGraph
 }


 Write-Host Connecting to Microsoft Graph...
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
 }
 else
 {
  Connect-MgGraph -Scopes "Application.Read.All"  -NoWelcome
 }
}

Connect_MgGraph
$Location=Get-Location
$ExportCSV = "$Location\EntraId_AppRegistration_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult=""   
$ExportResults=@() 
$AppCount=0
$PrintedCount=0

if(($CertificatesOnly.IsPresent) -or ($ClientSecretsOnly.IsPresent) -or ($ExpiredOnly.IsPresent) -or ($ActiveOnly.IsPresent) -or ($SoonToExpireInDays -ne ""))
{
 $SwitchPresent=$True
}
else
{
 $SwitchPresent=$false
}
$RequiredProperties=@('DisplayName','AppId','Id','KeyCredentials','PasswordCredentials','CreatedDateTime','SigninAudience')
Get-MgApplication -All -Property $RequiredProperties | foreach {
 $AppCount++
 $AppName=$_.DisplayName
 Write-Progress -Activity "`n     Processed App registration: $AppCount - $AppName "
 $AppId=$_.Id
 $Secrets=$_.PasswordCredentials
 $Certificates=$_.KeyCredentials
 $AppCreationDate=$_.CreatedDateTime

 $SigninAudience=$_.SignInAudience
 $Owners=(Get-MgApplicationOwner -ApplicationId $AppId).AdditionalProperties.userPrincipalName
 $Owners=$Owners -join ","
 if($owners -eq "")
 {
  $Owners="-"
 }
 #Process through Secret keys

 if(!($CertificatesOnly.IsPresent) -or ($SwitchPresent -eq $false))
 {
  foreach($Secret in $Secrets )
  {
   $CredentialType="Client Secret"
   $Print=1
   $DisplayName=$Secret.DisplayName
   $Id=$Secret.KeyId
   $CreatedTime=$Secret.StartDateTime
   $ExpiryDate=$Secret.EndDateTime
   $ExpiryStatusCalculation=(New-TimeSpan -Start (Get-Date).Date -End $ExpiryDate).Days
  
   if($ExpiryStatusCalculation -lt 0)
   {
    $ExpiryStatus="Expired"
    $ExpiryDateCalculation=$ExpiryStatusCalculation * (-1)
    $FriendlyExpiryTime="Expired $ExpiryDateCalculation days ago"
   }
   else
   {
    $ExpiryStatus="Active"
    $FriendlyExpiryTime="Expires in $ExpiryStatusCalculation days"
   }
  
   #Filter for expired/active client secrets

   if(($ExpiredOnly.IsPresent) -and ($ExpiryStatus -eq "Active"))
   { 
    $Print=0
   }
   if($ActiveOnly.IsPresent -and $ExpiryStatus -eq "Expired")
   { 
    $Print=0
   }

   #Filter for soon-to-expire client secrets
   if(($SoonToExpireInDays -ne "") -and (($SoonToExpireInDays -lt $ExpiryStatusCalculation) -or ($ExpiryStatus -eq "Expired")))
   {
    $Print=0
   }

   if($Print -eq 1)
   {
   $PrintedCount++
   $ExportResult=[PSCustomObject]@{'App Name'=$AppName;'App Id'=$AppId;'App Owners'=$Owners;'App Creation Time'=$AppCreationDate;'Credential Type'=$CredentialType;'Name'=$DisplayName;'Id'=$Id;'Expiry Status'=$ExpiryStatus;'Expiry Date'=$ExpiryDate;'Days since/to Expiry'=$ExpiryStatusCalculation;'Friendly Expiry Date'=$FriendlyExpiryTime;'Creation Time'=$CreatedTime}
   $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
  }
  }
 }

 if(!($ClientSecretsOnly.IsPresent) -or ($SwitchPresent -eq $false))
 {
  foreach ($Certificate in $Certificates)
  {
   $CredentialType="Certificate"
   $Print=1
   $DisplayName=$Certificate.DisplayName
   $Id=$Certificate.KeyId
   $CreatedTime=$Certificate.StartDateTime
   $ExpiryDate=$Certificate.EndDateTime
   $ExpiryStatusCalculation=(New-TimeSpan -Start (Get-Date).Date -End $ExpiryDate).Days
   if($ExpiryStatusCalculation -lt 0)
   {
    $ExpiryStatus="Expired"
    $ExpiryDateCalculation=$ExpiryStatusCalculation * (-1)
    $FriendlyExpiryTime="Expired $ExpiryDateCalculation days ago"
   }
   else
   {
    $ExpiryStatus="Active"
    $FriendlyExpiryTime="Expires in $ExpiryStatusCalculation days"
   }
   
   #Filter for expired/active certificates
   if(($ExpiredOnly.IsPresent) -and ($ExpiryStatus -eq "Active"))
   { 
    $Print=0
   }
   if($ActiveOnly.IsPresent -and $ExpiryStatus -eq "Expired")
   { 
    $Print=0
   }
   
   #Filter for soon-to-expire certificates
   if(($SoonToExpireInDays -ne "") -and (($SoonToExpireInDays -lt $ExpiryStatusCalculation) -or ($ExpiryStatus -eq "Expired")))
   { 
    $Print=0
   }

   if($Print -eq 1)
   {
   $PrintedCount++
   $ExportResult=[PSCustomObject]@{'App Name'=$AppName;'App Id'=$AppId;'App Owners'=$Owners;'App Creation Time'=$AppCreationDate;'Credential Type'=$CredentialType;'Name'=$DisplayName;'Id'=$Id;'Expiry Status'=$ExpiryStatus;'Expiry Date'=$ExpiryDate;'Days since/to Expiry'=$ExpiryStatusCalculation;'Friendly Expiry Date'=$FriendlyExpiryTime;'Creation Time'=$CreatedTime}
   $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
  }
  }
 }

 If(($Secrets.Count -eq 0) -and ($Certificates.Count -eq 0) -and ($SwitchPresent -eq $false))
 {
  $PrintedCount++
  $ExportResult=[PSCustomObject]@{'App Name'=$AppName;'App Id'=$AppId;'App Owners'=$Owners;'App Creation Time'=$AppCreationDate;'Credential Type'="-";'Name'="-";'Id'=$Id;'Expiry Status'="-";'Expiry Date'="-";'Days since/to Expiry'="-";'Friendly Expiry Date'="-";'Creation Time'="-"}
  $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n

 #Open output file after execution 
 if((Test-Path -Path $ExportCSV) -eq "True") 
 {
  Write-Host `nThe script processed $AppCount app registrations and the output file contains $PrintedCount records.
  Write-Host `n The Output file available in: -NoNewline -ForegroundColor Yellow
  Write-Host $ExportCSV 
   $Prompt = New-Object -ComObject wscript.shell      
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
   {   
    Invoke-Item "$ExportCSV"   
   } 
  }
 else
 {
  Write-Host No data found for the given criteria
 }
 