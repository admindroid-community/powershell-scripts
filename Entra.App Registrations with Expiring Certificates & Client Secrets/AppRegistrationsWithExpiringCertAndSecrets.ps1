<#
=============================================================================================
Name:           Retrieve Entra App registrations with expiring client secrets & certificates
Description:    The script retrieves the expiration date of app registrations' client secrets and certificates in Microsoft 365.
Version:        1.0
Website:        admindroid.com

Script Highlights:
1. The script automatically verifies and installs the Microsoft Graph PowerShell SDK module (if not installed already) upon your confirmation.
2. Exports all the Entra ID apps with expiring client secrets and certificates into a CSV file.
3. Allows to retrieve app registrations and their client secret expiry details.
4. Allows to retrieve app regsitrations and their certificate expiration details.
4. Allows to export the list soon to expire client secrets and certificates (i.e., 30 days, 90 days, etc)
5. Supports certificate-based authentication (CBA) too.
6. The script is scheduler-friendly.


For detailed Script execution: https://blog.admindroid.com/retrieve-entra-app-registrations-with-expiring-client-secrets-and-certificates/

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
    [int]$SoonToExpireInDays
)
Function Connect_MgGraph
{
 #Check for module installation
 $Module=Get-Module -Name microsoft.graph.beta -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Microsoft Graph PowerShell SDK is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Microsoft Graph PowerShell module..."
   Install-Module Microsoft.Graph.beta -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
  }
  else
  {
   Write-Host "Microsoft Graph Beta PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph cmdlet." 
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
$ExportCSV = "$Location\AppRegistration_with_Expiring_CertificatesAndSecrets_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult=""   
$ExportResults=@() 
$AppCount=0
$PrintedCount=0

if(($CertificatesOnly.IsPresent) -or ($ClientSecretsOnly.IsPresent) -or ($SoonToExpireInDays -ne ""))
{
 $SwitchPresent=$True
}
else
{
 $SwitchPresent=$false
}
$RequiredProperties=@('DisplayName','AppId','Id','KeyCredentials','PasswordCredentials','CreatedDateTime','SigninAudience')
Get-MgBetaApplication -All -Property $RequiredProperties | foreach {
 $AppCount++
 $AppName=$_.DisplayName
 Write-Progress -Activity "`n     Processed App registration: $AppCount - $AppName "
 $AppId=$_.Id
 $Secrets=$_.PasswordCredentials
 $Certificates=$_.KeyCredentials
 $AppCreationDate=$_.CreatedDateTime

 $SigninAudience=$_.SignInAudience
 $Owners=(Get-MgBetaApplicationOwner -ApplicationId $AppId).AdditionalProperties.userPrincipalName
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
    $Print=0
   }
   else
   {
    $ExpiryStatus="Active"
    $FriendlyExpiryTime="Expires in $ExpiryStatusCalculation days"
   }
  
   #Filter for soon-to-expire client secrets
   if(($SoonToExpireInDays -ne "") -and (($SoonToExpireInDays -lt $ExpiryStatusCalculation)))
   {
    $Print=0
   }

   if($Print -eq 1)
   {
   $PrintedCount++
   $ExportResult=[PSCustomObject]@{'App Name'=$AppName;'App Owners'=$Owners;'App Creation Time'=$AppCreationDate;'Credential Type'=$CredentialType;'Name'=$DisplayName;'Id'=$Id;'Creation Time'=$CreatedTime;'Expiry Date'=$ExpiryDate;'Days to Expiry'=$ExpiryStatusCalculation;'Friendly Expiry Date'=$FriendlyExpiryTime;'App Id'=$AppId}
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
    $Print=0
   }
   else
   {
    $ExpiryStatus="Active"
    $FriendlyExpiryTime="Expires in $ExpiryStatusCalculation days"
   }
   
   
   #Filter for soon-to-expire certificates
   if(($SoonToExpireInDays -ne "") -and (($SoonToExpireInDays -lt $ExpiryStatusCalculation)))
   { 
    $Print=0
   }

   if($Print -eq 1)
   {
   $PrintedCount++
   $ExportResult=[PSCustomObject]@{'App Name'=$AppName;'App Owners'=$Owners;'App Creation Time'=$AppCreationDate;'Credential Type'=$CredentialType;'Name'=$DisplayName;'Id'=$Id;'Creation Time'=$CreatedTime;'Expiry Date'=$ExpiryDate;'Days to Expiry'=$ExpiryStatusCalculation;'Friendly Expiry Date'=$FriendlyExpiryTime;'App Id'=$AppId}
   $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
  }
  }
 }
}


 #Open output file after execution 
 If($PrintedCount -eq 0)
 {
  Write-Host No data found for the given criteria
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 
 }
 else
 {
  Write-Host `nThe script processed $AppCount app registrations and the output file contains $PrintedCount records.
  if((Test-Path -Path $ExportCSV) -eq "True") 
  {

   Write-Host `n The Output file available in: -NoNewline -ForegroundColor Yellow
   Write-Host $ExportCSV 
   Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
   Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 
   $Prompt = New-Object -ComObject wscript.shell      
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
   {   
    Invoke-Item "$ExportCSV"   
   } 
  }
 }



 