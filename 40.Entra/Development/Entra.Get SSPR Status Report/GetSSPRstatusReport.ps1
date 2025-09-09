<#
=============================================================================================
Name:           Export Microsoft 365 Users' Self-service Password Reset (SSPR) Status Reports 
Description:    The script exports users' Self-service password reset status reports to CSV. 
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~  
1. The script exports 10 SSPR status reports. 
2. Exports SSPR status for Microsoft 365 users. 
3. Generates report on SSPR enabled users. 
4. Finds SSPR disabled users. 
5. Identifies users who are eligible but not registered for SSPR.  
6. Finds SSPR status for Microsoft 365 admins. 
7. Determines the SSPR status specifically for licensed users. 
8. The script can be executed with MFA-enabled accounts.  
9. It exports results to a CSV file for convenient data handling. 
10. The script installs the required Microsoft Graph Beta module upon user confirmation if not already installed. 
11. Supports certificate-based authentication (scheduler-friendly) method. 


For detailed Script execution: https://o365reports.com/2024/02/13/export-microsoft-365-users-self-service-password-reset-sspr-status-reports
============================================================================================
#>
Param
(

    [switch]$AdminsOnly,
    [switch]$LicensedUsersOnly,
    [switch]$SsprTurnedOnButUserNotRegistered,
    [Switch]$SsprEnabledUsers,
    [Switch]$SsprDisabledUsers,
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
$ExportCSV="$Location\SSPR_Status_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$Result=""   
$Results=@()  
$OutputCount=0
$ProcessedUsersCount=0
Write-Host "Generating M365 users' SSPR status report..." -ForegroundColor Cyan
Get-MgBetaReportAuthenticationMethodUserRegistrationDetail -All | ? { $_.UserType -eq 'member' } | foreach {
 $UPN=$_.UserPrincipalName
 $DisplayName=$_.UserDisplayName
 $IsAdmin=$_.IsAdmin
 $IsSsprEnabled=$_.IsSsprEnabled
 $IsSsprRegistered=$_.IsSsprRegistered
 $IsSsprCapable=$_.IsSsprCapable
 $RegisteredMethods=$_.MethodsRegistered
 $RegisteredMethods=$RegisteredMethods -join ","
 $UserPreferredAuthMethod=$_.UserPreferredMethodForSecondaryAuthentication
 $Print=1
 $ProcessedUsersCount++
 Write-Progress -Activity "`n     Processed user count: $ProcessedUsersCount "`n"  Currently Processing: $DisplayName"
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
 if($LicensedUsersOnly.IsPresent -and ($IsLicensed -ne "Licensed"))
 {
  $Print=0
 }
 if($AdminsOnly.IsPresent -and ($IsAdmin -ne $true))
 {
  $Print=0
 }
 if($SsprEnabledUsers.IsPresent -and ($IsSsprCapable -ne $true))
 {
  $Print=0
 }
 if($SsprDisabledUsers.IsPresent -and ($IsSsprCapable -ne $false))
 {
  $Print=0
 }
 if($SsprTurnedOnButUserNotRegistered.IsPresent -and ($IsSsprEnabled -eq $IsSsprRegistered))
 {
  $Print=0 
 }
 #Export result to csv
 if($Print -eq 1)
 {
  $OutputCount++
  $Result=@{'User Name'=$DisplayName;'UPN'=$upn;'Is SSPR Registered by User'=$IsSsprRegistered;'Is SSPR Enabled by Admins'=$IsSsprEnabled;'Department'=$Department;'Job Title'=$JobTitle;'License Status'=$IsLicensed;'Signin Enabled'=$SignInEnabled;'Is Admin'=$IsAdmin;'Registered Auth Methods'=$RegisteredMethods;'Default Auth Method'=$UserPreferredAuthMethod}
  $Results= New-Object PSObject -Property $Result  
  $Results | Select-Object 'User Name','UPN','Is SSPR Registered by User','Is SSPR Enabled by Admins','Department','Registered Auth Methods','Default Auth Method','Job Title','License Status','Signin Enabled','Is Admin'| Export-Csv -Path $ExportCSV -Notype -Append 
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
 
