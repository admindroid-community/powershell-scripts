<#
=============================================================================================
Name:           Export Microsoft 365 Users' SMS and Voice Authentication Retirement Report 
Description:    The script exports users affected by the retirement of Microsoft-provided SMS and voice authentication to CSV. 
Version:        1.0
Website:        https://blog.admindroid.com

Script Highlights: 
~~~~~~~~~~~~~~~~~  
1. Finds users who have SMS or voice (phone) authentication methods registered.
2. Rates the retirement's impact on each user's MFA and SSPR.
3. Shows the default MFA method that applies, including system-preferred MFA.
4. Filters users by MFA and SSPR impact levels.
5. Finds affected Microsoft 365 admins.
6. Finds affected users who are enabled for SSPR.
7. Finds affected licensed users.
8. The script can be executed with MFA-enabled accounts.  
9. It exports results to a CSV file for convenient data handling. 
10. The script installs the required Microsoft Graph module upon user confirmation if not already installed. 
11. Supports certificate-based authentication (scheduler-friendly) method. 


MFA Impact:
~~~~~~~~~~~
    High   - No MFA method other than phone. User must register a passkey at sign-in after the retirement.
    Medium - Another MFA method is registered, but SMS/voice is the default MFA method.
    Low    - Another MFA method is registered and is the default.

SSPR Impact:
~~~~~~~~~~~~
    High        - No SSPR method other than phone (fewer than 2 for admins). User can't reset password after the retirement.
    Medium      - One SSPR method other than phone. Enough only if the SSPR policy requires one method.
    Low         - Two SSPR methods other than phone.
    Not Enabled - User is not enabled for SSPR.
    Note: Email and Authenticator/OATH count as SSPR methods. Authenticator and OATH tokens together count as one method.
          Passkeys and security questions are not counted.

Note: The ratings use registered methods only. Users moved to a telephony provider through Microsoft Security Store are not affected.


For detailed Script execution: https://blog.admindroid.com/microsoft-entra-id-sms-voice-retirement-for-sspr
============================================================================================
#>
Param
(

    [switch]$AdminsOnly,
    [switch]$LicensedUsersOnly,
    [Switch]$SsprEnabledUsers,
    [ValidateSet('High','Medium','Low')]
    [string[]]$MfaImpact,
    [ValidateSet('High','Medium','Low')]
    [string[]]$SsprImpact,
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)



Function Connect_MgGraph
{
 $MsGraphModule =  Get-Module Microsoft.Graph -ListAvailable
 if($MsGraphModule -eq $null)
 { 
    Write-host "Important: Microsoft Graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
    $confirm = Read-Host Are you sure you want to install Microsoft Graph module? [Y] Yes [N] No  
    if($confirm -match "[yY]") 
    { 
        Write-host "Installing Microsoft Graph module..."
        Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber
        Write-host "Microsoft Graph module is installed in the machine successfully" -ForegroundColor Magenta 
        #importing required modules
        Import-Module Microsoft.Graph.Authentication
        Import-Module Microsoft.Graph.Reports
        Import-Module Microsoft.Graph.Users

    } 
    else
    { 
        Write-host "Exiting. `nNote: Microsoft Graph module must be available in your system to run the script" -ForegroundColor Red
        Exit 
    } 
 }
 #Disconnect Existing MgGraph session
 if($CreateSession.IsPresent)
 {
  Disconnect-MgGraph
 }
 #Connecting to MgGraph
 Write-Host Connecting to Microsoft Graph...
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint 
 }
 else
 {
  Connect-MgGraph -Scopes "User.Read.All","AuditLog.read.All"  -NoWelcome
 }
}
Connect_MgGraph

$Location=Get-Location
$ExportCSV="$Location\SMS_Voice_Retirement_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$Result=""   
$Results=@()  
$OutputCount=0
$ProcessedUsersCount=0
#Phone methods affected by the retirement
$PhoneTypes='mobilePhone','alternateMobilePhone','officePhone'
#Department, job title and licenses of all users
$UserLookup=@{}
Get-MgUser -All -Property Id,Department,JobTitle,AssignedLicenses | foreach { $UserLookup[$_.Id]=$_ }
Write-Host "Generating M365 users' SMS/voice retirement report..." -ForegroundColor Cyan
Get-MgReportAuthenticationMethodUserRegistrationDetail -All | ? { $_.UserType -eq 'member' } | foreach {
 $UPN=$_.UserPrincipalName
 $DisplayName=$_.UserDisplayName
 $IsAdmin=$_.IsAdmin
 $IsSsprEnabled=$_.IsSsprEnabled
 $IsSsprRegistered=$_.IsSsprRegistered
 $RegisteredMethods=$_.MethodsRegistered
 #Default MFA method that applies
 if($_.IsSystemPreferredAuthenticationMethodEnabled -and $_.SystemPreferredAuthenticationMethods)
 {
  $DefaultAuthMethod=@($_.SystemPreferredAuthenticationMethods)[0]
 }
 else
 {
  $DefaultAuthMethod=$_.UserPreferredMethodForSecondaryAuthentication
 }
 #Rate MFA impact
 $SmsVoiceMethods=@($RegisteredMethods | ? { $_ -in $PhoneTypes })
 $MfaOtherMethods=@($RegisteredMethods | ? { $_ -and ($_ -notin $PhoneTypes) -and ($_ -notin 'email','securityQuestion','temporaryAccessPass') })
 if($SmsVoiceMethods.Count -eq 0) { $MfaImpactLevel='None' }
 elseif($MfaOtherMethods.Count -eq 0) { $MfaImpactLevel='High' }
 elseif($DefaultAuthMethod -in 'sms','voiceMobile','voiceAlternateMobile','voiceOffice') { $MfaImpactLevel='Medium' }
 else { $MfaImpactLevel='Low' }
 #Rate SSPR impact
 $SsprMethodCount=0
 if('email' -in $RegisteredMethods) { $SsprMethodCount++ }
 if($RegisteredMethods | ? { $_ -in 'microsoftAuthenticatorPush','softwareOneTimePasscode','hardwareOneTimePasscode' }) { $SsprMethodCount++ }
 if($SmsVoiceMethods.Count -eq 0) { $SsprImpactLevel='None' }
 elseif($IsSsprEnabled -ne $true) { $SsprImpactLevel='Not Enabled' }
 elseif(($SsprMethodCount -eq 0) -or ($IsAdmin -and ($SsprMethodCount -lt 2))) { $SsprImpactLevel='High' }
 elseif($SsprMethodCount -eq 1) { $SsprImpactLevel='Medium' }
 else { $SsprImpactLevel='Low' }
 $RegisteredMethods=$RegisteredMethods -join ","
 $Print=1
 $ProcessedUsersCount++
 Write-Progress -Activity "`n     Processed user count: $ProcessedUsersCount "`n"  Currently Processing: $DisplayName"
 $UserDetails=$UserLookup[$_.Id]
 if($UserDetails.AssignedLicenses.Count -gt 0)
 {
  $IsLicensed="Licensed"
 }
 else
 {
  $IsLicensed="Unlicensed"
 }
 $Department=$UserDetails.Department
 $JobTitle=$UserDetails.JobTitle
 if($LicensedUsersOnly.IsPresent -and ($IsLicensed -ne "Licensed"))
 {
  $Print=0
 }
 if($AdminsOnly.IsPresent -and ($IsAdmin -ne $true))
 {
  $Print=0
 }
 if($SsprEnabledUsers.IsPresent -and ($IsSsprEnabled -ne $true))
 {
  $Print=0
 }
 if($MfaImpactLevel -eq 'None')
 {
  $Print=0
 }
 if($MfaImpact -and ($MfaImpactLevel -notin $MfaImpact))
 {
  $Print=0
 }
 if($SsprImpact -and ($SsprImpactLevel -notin $SsprImpact))
 {
  $Print=0
 }
 #Export result to csv
 if($Print -eq 1)
 {
  $OutputCount++
  $Result=@{'User Name'=$DisplayName;'UPN'=$upn;'Is SSPR Registered by User'=$IsSsprRegistered;'Is SSPR Enabled by Admins'=$IsSsprEnabled;'Department'=$Department;'Job Title'=$JobTitle;'License Status'=$IsLicensed;'Is Admin'=$IsAdmin;'Registered Auth Methods'=$RegisteredMethods;'Default MFA Method'=$DefaultAuthMethod;'SMS/Voice Methods'=($SmsVoiceMethods -join ",");'MFA Impact'=$MfaImpactLevel;'SSPR Impact'=$SsprImpactLevel}
  $Results= New-Object PSObject -Property $Result  
  $Results | Select-Object 'User Name','UPN','MFA Impact','SSPR Impact','SMS/Voice Methods','Registered Auth Methods','Default MFA Method','Is SSPR Enabled by Admins','Is SSPR Registered by User','Department','Job Title','License Status','Is Admin'| Export-Csv -Path $ExportCSV -Notype -Append 
 }
}

 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 3500+ Microsoft 365 reports and 450+ management actions. ~~" -ForegroundColor Green

 #Open output file after execution 
 If($OutputCount -eq 0)
 {
  Write-Host No data found for the given criteria
 }
 else
 {
  Write-Host `nThe output file contains $OutputCount accounts.
  if((Test-Path -Path $ExportCSV) -eq "True") 
  {

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
 }
 
