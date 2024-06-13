<#
=============================================================================================
Name:           Export MFA status  Microsoft Graph PowerShell
Description:    This script exports Microsoft 365 users' per-user MFA status report to CSV file
Version:        1.0
Website:        blog.admindroid.com


Script Highlights :
~~~~~~~~~~~~~~~~~

1.Generates 5+ MFA status reports 
2.Helps to filter and find MFA enabled users 
3.Identifies MFA disabled users 
4.Helps to track MFA enforced users 
5.Exports MFA status report for licensed users alone 
6.Checks MFA status for sign-in allowed users alone 
7.Helps to generate more granular MFA reports by combining multiple filtering params. 
8.The script uses MS Graph PowerShell and installs MS Graph Beta PowerShell SDK (if not installed already) upon your confirmation. 
9.The script can be executed with an MFA enabled account too. 
10.Exports report results as a CSV file. 
11.The script is scheduler friendly. 
12.It can be executed with certificate-based authentication (CBA) too.

For detailed script execution: https://blog.admindroid.com/export-mfa-status-report-for-entra-id-accounts-using-powershell
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$CreateSession,
    [switch]$MFAEnabledUsersOnly,
    [switch]$MFADisabledUsersOnly,
    [switch]$MFAEnforcedUsersOnly,
    [switch]$LicensedUsersOnly,
    [switch]$SignInAllowedUsersOnly,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint

)
Function Connect_MgGraph
{
 #Check for module installation
 $MsGraphBetaModule =  Get-Module Microsoft.Graph.Beta -ListAvailable
 if($MsGraphBetaModule -eq $null)
 { 
    Write-host "Important: Microsoft Graph Beta module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
    $confirm = Read-Host Are you sure you want to install Microsoft Graph Beta module? [Y] Yes [N] No  
    if($confirm -match "[Y]") 
    { 
        Write-host "Installing Microsoft Graph Beta module..."
        Install-Module Microsoft.Graph.Beta -Scope CurrentUser -AllowClobber
        Write-host "Microsoft Graph Beta module is installed in the machine successfully" -ForegroundColor Magenta 
    } 
    else
    { 
        Write-host "Exiting. `nNote: Microsoft Graph Beta module must be available in your system to run the script" -ForegroundColor Red
        Exit 
    } 
 }
 #Disconnect Existing MgGraph session
 if($CreateSession.IsPresent)
 {
  Disconnect-MgGraph
 }
 #Connecting to MgGraph beta
 Write-Host Connecting to Microsoft Graph...
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
 }
 else
 {
  Connect-MgGraph -Scopes "User.Read.All","UserAuthenticationMethod.Read.All","Policy.ReadWrite.AuthenticationMethod" -NoWelcome
  }
}
Connect_MgGraph


$ProcessedUserCount=0
$ExportCount=0
 #Set output file 
 $Location=Get-Location
 $ExportCSV="$Location\MfaStatusReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
  $Result=""  
 $Results=@()

#Get all users
Get-MgBetaUser -All | foreach {
 $ProcessedUserCount++
 $Name= $_.DisplayName
 $UPN=$_.UserPrincipalName
 $Department=$_.Department
 $UserId=$_.Id
 if($_.AccountEnabled -eq $true)
 {
  $SigninStatus="Allowed"
 }
 else
 {
  $SigninStatus="Blocked"
 }
 if(($_.AssignedLicenses).Count -ne 0)
 {
  $LicenseStatus="Licensed"
 }
 else
 {
  $LicenseStatus="Unlicensed"
 }   
 $Is3rdPartyAuthenticatorUsed="False"
 $MFAPhone="-"
 $MicrosoftAuthenticatorDevice="-"
 Write-Progress -Activity "`n     Processed users count: $ProcessedUserCount "`n"  Currently processing user: $Name"
 [array]$MFAData=Get-MgBetaUserAuthenticationMethod -UserId $UPN
 $AuthenticationMethod=@()
 $AdditionalDetails=@()
 
 foreach($MFA in $MFAData)
 { 
   Switch ($MFA.AdditionalProperties["@odata.type"]) 
   { 
    "#microsoft.graph.passwordAuthenticationMethod"
    {
     $AuthMethod     = 'PasswordAuthentication'
     $AuthMethodDetails = $MFA.AdditionalProperties["displayName"] 
    } 
    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"  
    { # Microsoft Authenticator App
     $AuthMethod     = 'AuthenticatorApp'
     $AuthMethodDetails = $MFA.AdditionalProperties["displayName"] 
     $MicrosoftAuthenticatorDevice=$MFA.AdditionalProperties["displayName"]
    }
    "#microsoft.graph.phoneAuthenticationMethod"                  
    { # Phone authentication
     $AuthMethod     = 'PhoneAuthentication'
     $AuthMethodDetails = $MFA.AdditionalProperties["phoneType", "phoneNumber"] -join ' ' 
     $MFAPhone=$MFA.AdditionalProperties["phoneNumber"]
    } 
    "#microsoft.graph.fido2AuthenticationMethod"                   
    { # FIDO2 key
     $AuthMethod     = 'Passkeys(FIDO2)'
     $AuthMethodDetails = $MFA.AdditionalProperties["model"] 
    }  
    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" 
    { # Windows Hello
     $AuthMethod     = 'WindowsHelloForBusiness'
     $AuthMethodDetails = $MFA.AdditionalProperties["displayName"] 
    }                        
    "#microsoft.graph.emailAuthenticationMethod"        
    { # Email Authentication
     $AuthMethod     = 'EmailAuthentication'
     $AuthMethodDetails = $MFA.AdditionalProperties["emailAddress"] 
    }               
    "microsoft.graph.temporaryAccessPassAuthenticationMethod"   
    { # Temporary Access pass
     $AuthMethod     = 'TemporaryAccessPass'
     $AuthMethodDetails = 'Access pass lifetime (minutes): ' + $MFA.AdditionalProperties["lifetimeInMinutes"] 
    }
    "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" 
    { # Passwordless
     $AuthMethod     = 'PasswordlessMSAuthenticator'
     $AuthMethodDetails = $MFA.AdditionalProperties["displayName"] 
    }      
    "#microsoft.graph.softwareOathAuthenticationMethod"
    { 
      $AuthMethod     = 'SoftwareOath'
      $Is3rdPartyAuthenticatorUsed="True"            
    }
    
   }
   $AuthenticationMethod +=$AuthMethod
   if($AuthMethodDetails -ne $null)
   {
    $AdditionalDetails +="$AuthMethod : $AuthMethodDetails"
   }
  }
  #To remove duplicate authentication methods
  $AuthenticationMethod =$AuthenticationMethod | Sort-Object | Get-Unique
  $AuthenticationMethods= $AuthenticationMethod  -join ","
  $AdditionalDetail=$AdditionalDetails -join ", "
  $Print=1
  #Per-user MFA status
  $MFAStatus=(Invoke-MgGraphRequest -Method GET -Uri "/beta/users/$UserId/authentication/requirements").perUserMfaState
 <# [array]$StrongMFAMethods=("Fido2","PhoneAuthentication","PasswordlessMSAuthenticator","AuthenticatorApp","WindowsHelloForBusiness")
  $MFAStatus="Disabled"
 

  foreach($StrongMFAMethod in $StrongMFAMethods)
  {
   if($AuthenticationMethod -contains $StrongMFAMethod)
   {
    $MFAStatus="Strong"
    break
   }
  }

  if(($MFAStatus -ne "Strong") -and ($AuthenticationMethod -contains "SoftwareOath"))
  {
   $MFAStatus="Weak"
  } #>

  #Filter result based on MFA status
  if($MFADisabledUsersOnly.IsPresent -and $MFAStatus -ne "disabled")
  {
   $Print=0
  }
  if($MFAEnabledUsersOnly.IsPresent -and $MFAStatus -ne "enabled")
  {
   $Print=0
  }
  if($MFAEnforcedUsersOnly.IsPresent -and $MFAStatus -ne "enforced")
  {
   $Print=0
  }

  #Filter result based on license status
  if($LicensedUsersOnly.IsPresent -and ($LicenseStatus -eq "Unlicensed"))
  {
   $Print=0
  }

  #Filter result based on signin status
  if($SignInAllowedUsersOnly.IsPresent -and ($SigninStatus -eq "Blocked"))
  {
   $Print=0
  }
 
 if($Print -eq 1)
 {
  $ExportCount++
  $Result=@{'Name'=$Name;'UPN'=$UPN;'Department'=$Department;'License Status'=$LicenseStatus;'SignIn Status'=$SigninStatus;'Registered Authentication Methods'=$AuthenticationMethods;'Per-user MFA Status'=$MFAStatus;'MFA Phone'=$MFAPhone;'Microsoft Authenticator Configured Device'=$MicrosoftAuthenticatorDevice;'Is 3rd-Party Authenticator Used'=$Is3rdPartyAuthenticatorUsed;'Additional Details'=$AdditionalDetail} 
  $Results= New-Object PSObject -Property $Result 
  $Results | Select-Object Name,UPN,'Per-user MFA Status',Department,'License Status','SignIn Status','Registered Authentication Methods','MFA Phone','Microsoft Authenticator Configured Device','Is 3rd-Party Authenticator Used','Additional Details' | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

if((Test-Path -Path $ExportCSV) -eq "True") 
 {
  Write-Host `nThe exported report contains $ExportCount users.
  Write-Host `nPer-user MFA status report available in: -NoNewline -Foregroundcolor Yellow; Write-Host $ExportCSV
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
 else
 {
  Write-Host No users found.
 }
 