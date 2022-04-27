
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$CreateSession,
    [switch]$MFAEnabled,
    [switch]$MFADisabled,
    [switch]$LicensedUsersOnly,
    [switch]$SignInAllowedUsersOnly

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
 #Connecting to MgGraph beta
 Select-MgProfile -Name beta
 Write-Host Connecting to Microsoft Graph...
 Connect-MgGraph -Scopes "User.Read.All","UserAuthenticationMethod.Read.All"
}
Connect_MgGraph
if((Get-MgContext) -ne "")
{
 Write-Host Connected to Microsoft Graph PowerShell using (Get-MgContext).Account account -ForegroundColor Yellow
}
$ProcessedUserCount=0
$ExportCount=0
 #Set output file 
 $ExportCSV=".\MfaStatusReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
  $Result=""  
 $Results=@()

#Get all users
Get-MgUser -All -Filter "UserType eq 'Member'" | foreach {
 $ProcessedUserCount++
 $Name= $_.DisplayName
 $UPN=$_.UserPrincipalName
 $Department=$_.Department
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
 [array]$MFAData=Get-MgUserAuthenticationMethod -UserId $UPN
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
     $AuthMethod     = 'Fido2'
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
  #Determine MFA status
  [array]$StrongMFAMethods=("Fido2","PhoneAuthentication","PasswordlessMSAuthenticator","AuthenticatorApp","WindowsHelloForBusiness")
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
  }
  #Filter result based on MFA status
  if($MFADisabled.IsPresent -and $MFAStatus -ne "Disabled")
  {
   $Print=0
  }
  if($MFAEnabled.IsPresent -and $MFAStatus -eq "Disabled")
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
  $Result=@{'Name'=$Name;'UPN'=$UPN;'Department'=$Department;'License Status'=$LicenseStatus;'SignIn Status'=$SigninStatus;'Authentication Methods'=$AuthenticationMethods;'MFA Status'=$MFAStatus;'MFA Phone'=$MFAPhone;'Microsoft Authenticator Configured Device'=$MicrosoftAuthenticatorDevice;'Is 3rd-Party Authenticator Used'=$Is3rdPartyAuthenticatorUsed;'Additional Details'=$AdditionalDetail} 
  $Results= New-Object PSObject -Property $Result 
  $Results | Select-Object Name,UPN,Department,'License Status','SignIn Status','Authentication Methods','MFA Status','MFA Phone','Microsoft Authenticator Configured Device','Is 3rd-Party Authenticator Used','Additional Details' | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

if((Test-Path -Path $ExportCSV) -eq "True") 
 {
  Write-Host `nThe output file contains $ExportCount users.
  Write-Host `nThe Output file available in the current working directory with name: $ExportCSV -ForegroundColor Green
  Write-Host `n"For more Microsoft 365 PowerShell scripts, visit: https://o365reports.com"
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
  Write-Host No users found
 }

 <#
=============================================================================================
Name:           Export Office 365 users' MFA status using Microsoft Graph PowerShell
Website:        o365reports.com
For detailed script execution: https://o365reports.com/2022/04/27/get-mfa-status-of-office-365-users-using-microsoft-graph-powershell
============================================================================================
#>