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
 Write-Host "MFA Disabled user report available in: $ExportCSV"
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
 Write-Host "MFA Enabled user report available in: $ExportCSVReport"
 $Prompt = New-Object -ComObject wscript.shell
 $UserInput = $Prompt.popup("Do you want to open output file?",`
 0,"Open Output File",4)
 If ($UserInput -eq 6)
 {
  Invoke-Item "$ExportCSVReport"
 }
 Write-Host Exported report has $PrintedUser users
}
Else
{
  Write-Host No user found that matches your criteria.
}
#Clean up session
Get-PSSession | Remove-PSSession
