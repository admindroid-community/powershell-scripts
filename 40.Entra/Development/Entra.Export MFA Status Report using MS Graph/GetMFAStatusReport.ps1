<#
=============================================================================================
Name:           Export Office 365 users' MFA status using Microsoft Graph PowerShell
Description:    This script exports O365 users MFA status report to CSV file
Version:        1.0
Website:        o365reports.com
Script by:      O365Reports Team

Script Highlights :
~~~~~~~~~~~~~~~~~

1.	The script exports MFA status for all users. 
2.	You can filter results based on MFA status. I.e., you can export MFA enabled/disabled users separately. 
3.	Exports report to CSV file 
4.	You can filter the result to display Licensed users alone. 
5.	You can generate MFA report for sign-in allowed users only. 
6.	Shows MFA registration done through Conditional Access and Security Defaults too.
7.	Automatically installs Microsoft Graph PowerShell module (if not installed already) upon your confirmation. 


For detailed script execution: https://o365reports.com/2022/04/27/get-mfa-status-of-office-365-users-using-microsoft-graph-powershell
============================================================================================
#>
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
 $MsGraphBetaModule =  Get-Module Microsoft.Graph.Beta -ListAvailable
 if($MsGraphBetaModule -eq $null)
 { 
    Write-host "Important: Microsoft Graph Beta module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
    $confirm = Read-Host Are you sure you want to install Microsoft Graph Beta module? [Y] Yes [N] No  
    if($confirm -match "[yY]") 
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
 Connect-MgGraph -Scopes "User.Read.All","UserAuthenticationMethod.Read.All"
}
Connect_MgGraph
Write-Host "`nNote: If you encounter module related conflicts, run the script in a fresh PowerShell window.`n" -ForegroundColor Yellow
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
Get-MgBetaUser -All -Filter "UserType eq 'Member'" | foreach {
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
  Write-Host `nThe Output file available in the current working directory with name: -NoNewline -Foregroundcolor Yellow; Write-Host $ExportCSV
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
  Write-Host No users found.
 }
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n