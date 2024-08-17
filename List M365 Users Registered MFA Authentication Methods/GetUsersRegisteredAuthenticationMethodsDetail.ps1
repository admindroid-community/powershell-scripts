<#
=============================================================================================

Name         : Get Microsoft 365 Users' Registered MFA Methods With PowerShell  
Version      : 1.0
website      : o365reports.com

-----------------
Script Highlights
-----------------
1. This script exports registered authentication methods for all users. 
2. Allows to get users with system preferred MFA enabled/disabled details. 
3. It specifically identifies registered authentication methods for admins alone. 
4. Tracks MFA authentication registration method for licensed users only. 
5. Retrieves users with no registered authentication method. 
6. Automatically installs the module Microsoft Graph with your confirmation. 
7. The script can be executed with an MFA-enabled account too. 
8. Exports report results as a CSV file. 
9. The script is schedular-friendly, making it easy to automate. 
10. It supports certificate-based authentication (CBA) too. 

For detailed Script execution:  https://o365reports.com/2024/08/13/get-registered-mfa-methods-in-microsoft-365-with-powershell/
============================================================================================
\#>
Param
(

    [switch]$AdminsOnly,
    [switch]$LicensedUsersOnly,
    [switch]$RegisteredUsersOnly,
    [Switch]$UserswithNoRegistration,
    [Switch]$UsersWithSystemPreferredMFA,
    [Switch]$UsersWithoutSystemPreferredMFA,
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)



Function Connect_MgGraph
{
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
        #importing required modules
        Import-Module Microsoft.Graph.Authentication
        Import-Module Microsoft.Graph.Beta.Report

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
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint 
 }
 else
 {
  Connect-MgGraph -Scopes "User.Read.All","AuditLog.read.All"  -NoWelcome
 }
}
Connect_MgGraph

$Location=Get-Location
$ExportCSV="$Location\M365Users_RegisteredAuthenticationMethods_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv" 
$Result=""   
$Results=@()  
$OutputCount=0
$ProcessedUsersCount=0
Write-Host "Retrieving M365 users' registered authentication methods..." -ForegroundColor Cyan
Get-MgBetaReportAuthenticationMethodUserRegistrationDetail | ? { $_.UserType -eq 'member' } | foreach {
 $UPN=$_.UserPrincipalName
 $DisplayName=$_.UserDisplayName
 $IsAdmin=$_.IsAdmin
 $RegisteredMethods=$_.MethodsRegistered
 $MethodCount=($RegisteredMethods | Measure-Object).Count
 $RegisteredMethods=$RegisteredMethods -join ","
 $UserPreferredAuthMethod=$_.UserPreferredMethodForSecondaryAuthentication
 $IsSystemPreferredAuthenticationEnabled=$_.IsSystemPreferredAuthenticationMethodEnabled
 $SystemPreferredAuthenticationMethod=$_.SystemPreferredAuthenticationMethods
 $SystemPreferredAuthenticationMethod=$SystemPreferredAuthenticationMethod -join ","
 $Print=1
 $ProcessedUsersCount++
 Write-Progress -Activity "`n     Processed user count: $ProcessedUsersCount "`n"  Currently Processing: $DisplayName"
 if($UserPreferredAuthMethod -eq "none")
 {
  $UserPreferredAuthMethod="-"
 }
 $UserDetails=Get-MgBetaUser -UserId $UPN
 if($UserDetails.AssignedLicenses -ne "")
 {
  $IsLicensed="Licensed"
 }
 else
 {
  $IsLicensed="Unlicensed"
 }
 $SignInEnabled=$UserDetails.AccountEnabled
 $Department=$UserDetails.Department
 $JobTitle=$UserDetails.JobTitle

 #Filter for licensed users
 if($LicensedUsersOnly.IsPresent -and ($IsLicensed -ne "Licensed"))
 {
  $Print=0
 }

 #Filter for administrators
 if($AdminsOnly.IsPresent -and ($IsAdmin -ne $true))
 {
  $Print=0
 }

 #Filter users based on their system-preferred MFA status
 if($UsersWithSystemPreferredMFA.IsPresent -and $IsSystemPreferredAuthenticationEnabled -ne $true)
 {
  $Print=0
 }
 elseif($UsersWithoutSystemPreferredMFA.IsPresent -and $IsSystemPreferredAuthenticationEnabled -ne $false)
 {
  $Print=0
 }

 #Filter users based on their Authentication method regsitered state
 if($RegisteredUsersOnly.IsPresent -and $MethodCount -eq "0")
 {
  $Print=0
 }
 elseif($UserswithNoRegistration.IsPresent -and $MethodCount -ne "0")
 {
  $Print=0
 }


 if($IsSystemPreferredAuthenticationEnabled -eq $true)
 {
  $IsSystemPreferredAuthenticationEnabled = "Enabled"
 }
 else
 {
  $IsSystemPreferredAuthenticationEnabled = "Disabled"
 }

 #Export result to csv
 if($Print -eq 1)
 {
  $OutputCount++
  $Result=@{'User Name'=$DisplayName;'UPN'=$upn;'System Preferred MFA Status'=$IsSystemPreferredAuthenticationEnabled;'System Preferred MFA Method'=$SystemPreferredAuthenticationMethod;'Department'=$Department;'Job Title'=$JobTitle;'License Status'=$IsLicensed;'Signin Enabled'=$SignInEnabled;'Is Admin'=$IsAdmin;'Registered Auth Methods'=$RegisteredMethods;'Default Auth Method'=$UserPreferredAuthMethod}
  $Results= New-Object PSObject -Property $Result  
  $Results | Select-Object 'User Name','UPN','Registered Auth Methods','Default Auth Method','System Preferred MFA Status','System Preferred MFA Method','Department','Job Title','License Status','Signin Enabled','Is Admin'| Export-Csv -Path $ExportCSV -Notype -Append 
 }
}

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

 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 
