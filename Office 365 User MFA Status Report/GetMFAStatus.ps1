<#
=============================================================================================
Name:           Export Office 365 MFA status report
Description:    This script exports Microsoft 365 MFA status report to CSV
Version:        2.2
website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~

1.The result can be filtered based on MFA status. i.e., you can filter MFA enabled users/enforced users/disabled users alone. For example using the ‘EnabledOnly‘ flag you shall export Office 365 users’ MFA enabled status to CSV file.
2.Exports result to CSV file. 
3.Result can be filtered based on Admin users.
4.You can filter result to display Licensed users alone.
5.You can filter result based on SignIn Status (SignIn allowed/denied).
6.The script produces different output files based on MFA status. 
7.You can use this script to get users’ MFA status set by Conditional Access.
8.The script can be executed with MFA enabled account. 
9.Using the ‘Admin Roles’ column, you can find users with admin roles that are not protected with MFA. For example, you can find Global Admins without MFA.
10.The script is scheduler friendly. i.e., credentials can be passed as parameter instead of saving inside the script. 

For detailed Script execution: https://o365reports.com/2019/05/09/export-office-365-users-mfa-status-csv
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$DisabledOnly,
    [switch]$EnabledOnly,
    [switch]$EnforcedOnly,
    [switch]$ConditionalAccessOnly,
    [switch]$AdminOnly,
    [switch]$LicensedUserOnly,
    [Nullable[boolean]]$SignInAllowed = $null,
    [string]$UserName,
    [string]$Password
)
#Check for MSOnline module
$Modules=Get-Module -Name MSOnline -ListAvailable
if($Modules.count -eq 0)
{
  Write-Host  Please install MSOnline module using below command: `nInstall-Module MSOnline  -ForegroundColor yellow
  Exit
}

#Storing credential in script for scheduling purpose/ Passing credential as parameter
if(($UserName -ne "") -and ($Password -ne ""))
{
 $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
 $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
 Connect-MsolService -Credential $credential
}
else
{
 Connect-MsolService | Out-Null
}
$Result=""
$Results=@()
$UserCount=0
$PrintedUser=0

#Output file declaration
$ExportCSV=".\MFADisabledUserReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
$ExportCSVReport=".\MFAEnabledUserReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"


#Loop through each user
Get-MsolUser -All | foreach{
 $UserCount++
 $DisplayName=$_.DisplayName
 $Upn=$_.UserPrincipalName
 $MFAStatus=$_.StrongAuthenticationRequirements.State
 $MethodTypes=$_.StrongAuthenticationMethods
 $RolesAssigned=""
 Write-Progress -Activity "`n     Processed user count: $UserCount "`n"  Currently Processing: $DisplayName"
 if($_.BlockCredential -eq "True")
 {
  $SignInStatus="False"
  $SignInStat="Denied"
 }
 else
 {
  $SignInStatus="True"
  $SignInStat="Allowed"
 }

 #Filter result based on SignIn status
 if(($SignInAllowed -ne $null) -and ([string]$SignInAllowed -ne [string]$SignInStatus))
 {
  return
 }

 #Filter result based on License status
 if(($LicensedUserOnly.IsPresent) -and ($_.IsLicensed -eq $False))
 {
  return
 }

 if($_.IsLicensed -eq $true)
 {
  $LicenseStat="Licensed"
 }
 else
 {
  $LicenseStat="Unlicensed"
 }

 #Check for user's Admin role
 $Roles=(Get-MsolUserRole -UserPrincipalName $upn).Name
 if($Roles.count -eq 0)
 {
  $RolesAssigned="No roles"
  $IsAdmin="False"
 }
 else
 {
  $IsAdmin="True"
  foreach($Role in $Roles)
  {
   $RolesAssigned=$RolesAssigned+$Role
   if($Roles.indexof($role) -lt (($Roles.count)-1))
   {
    $RolesAssigned=$RolesAssigned+","
   }
  }
 }

 #Filter result based on Admin users
 if(($AdminOnly.IsPresent) -and ([string]$IsAdmin -eq "False"))
 {
  return
 }

 #Check for MFA enabled user
 if(($MethodTypes -ne $Null) -or ($MFAStatus -ne $Null) -and (-Not ($DisabledOnly.IsPresent) ))
 {
  #Check for Conditional Access
  if($MFAStatus -eq $null)
  {
   $MFAStatus='Enabled via Conditional Access'
  }

  #Filter result based on EnforcedOnly filter
  if((([string]$MFAStatus -eq "Enabled") -or ([string]$MFAStatus -eq "Enabled via Conditional Access")) -and ($EnforcedOnly.IsPresent))
  {
   return
  }

  #Filter result based on EnabledOnly filter
  if(([string]$MFAStatus -eq "Enforced") -and ($EnabledOnly.IsPresent))
  {
   return
  }

  #Filter result based on MFA enabled via Other source
  if((($MFAStatus -eq "Enabled") -or ($MFAStatus -eq "Enforced")) -and ($ConditionalAccessOnly.IsPresent))
  {
   return
  }

  $Methods=""
  $MethodTypes=""
  $MethodTypes=$_.StrongAuthenticationMethods.MethodType
  $DefaultMFAMethod=($_.StrongAuthenticationMethods | where{$_.IsDefault -eq "True"}).MethodType
  $MFAPhone=$_.StrongAuthenticationUserDetails.PhoneNumber
  $MFAEmail=$_.StrongAuthenticationUserDetails.Email

  if($MFAPhone -eq $Null)
  { $MFAPhone="-"}
  if($MFAEmail -eq $Null)
  { $MFAEmail="-"}

  if($MethodTypes -ne $Null)
  {
   $ActivationStatus="Yes"
   foreach($MethodType in $MethodTypes)
   {
    if($Methods -ne "")
    {
     $Methods=$Methods+","
    }
    $Methods=$Methods+$MethodType
   }
  }

  else
  {
   $ActivationStatus="No"
   $Methods="-"
   $DefaultMFAMethod="-"
   $MFAPhone="-"
   $MFAEmail="-"
  }

  #Print to output file
  $PrintedUser++
  $Result=@{'DisplayName'=$DisplayName;'UserPrincipalName'=$upn;'MFAStatus'=$MFAStatus;'ActivationStatus'=$ActivationStatus;'DefaultMFAMethod'=$DefaultMFAMethod;'AllMFAMethods'=$Methods;'MFAPhone'=$MFAPhone;'MFAEmail'=$MFAEmail;'LicenseStatus'=$LicenseStat;'IsAdmin'=$IsAdmin;'AdminRoles'=$RolesAssigned;'SignInStatus'=$SigninStat}
  $Results= New-Object PSObject -Property $Result
  $Results | Select-Object DisplayName,UserPrincipalName,MFAStatus,ActivationStatus,DefaultMFAMethod,AllMFAMethods,MFAPhone,MFAEmail,LicenseStatus,IsAdmin,AdminRoles,SignInStatus | Export-Csv -Path $ExportCSVReport -Notype -Append
 }

 #Check for MFA disabled user
 elseif(($DisabledOnly.IsPresent) -and ($MFAStatus -eq $Null) -and ($_.StrongAuthenticationMethods.MethodType -eq $Null))
 {
  $MFAStatus="Disabled"
  $Department=$_.Department
  if($Department -eq $Null)
  { $Department="-"}
  $PrintedUser++
  $Result=@{'DisplayName'=$DisplayName;'UserPrincipalName'=$upn;'Department'=$Department;'MFAStatus'=$MFAStatus;'LicenseStatus'=$LicenseStat;'IsAdmin'=$IsAdmin;'AdminRoles'=$RolesAssigned; 'SignInStatus'=$SigninStat}
  $Results= New-Object PSObject -Property $Result
  $Results | Select-Object DisplayName,UserPrincipalName,Department,MFAStatus,LicenseStatus,IsAdmin,AdminRoles,SignInStatus | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

#Open output file after execution
Write-Host `nScript executed successfully

if((Test-Path -Path $ExportCSV) -eq "True")
{
 Write-Host " MFA Disabled user report available in:" -NoNewline -ForegroundColor Yellow
 Write-Host $ExportCSV `n 
 $Prompt = New-Object -ComObject wscript.shell
 $UserInput = $Prompt.popup("Do you want to open output file?",`
 0,"Open Output File",4)
 If ($UserInput -eq 6)
 {
  Invoke-Item "$ExportCSV"
 }
 Write-Host Exported report has $PrintedUser users
}
elseif((Test-Path -Path $ExportCSVReport) -eq "True")
{
 Write-Host ""
 Write-Host " MFA Enabled user report available in:" -NoNewline -ForegroundColor Yellow
 Write-Host $ExportCSVReport `n
 $Prompt = New-Object -ComObject wscript.shell
 $UserInput = $Prompt.popup("Do you want to open output file?",`
 0,"Open Output File",4)
 If ($UserInput -eq 6)
 {
  Invoke-Item "$ExportCSVReport"
 }
 Write-Host Exported report has $PrintedUser users  
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n `n 
}
Else
{
  Write-Host No user found that matches your criteria. 
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n `n 
}
#Clean up session
Get-PSSession | Remove-PSSession
