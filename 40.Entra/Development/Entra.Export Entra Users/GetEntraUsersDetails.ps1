<#
=============================================================================================

Name         : Export all users in Microsoft 365 Using PowerShell  
Version      : 1.0
website      : o365reports.com

-----------------
Script Highlights
-----------------
1. The script automatically verifies and installs the Microsoft Graph PowerShell SDK module (if not installed already) upon your confirmation.
2. Exports all users from Microsoft Entra.
3. Allows filtering and exporting users that match the selected filters.
    -> Guest users
    -> Sign-in enabled users
    -> Sign-in blocked users
    -> License assigned users
    -> Users without any license
    -> Users without a manager
4. Identifies recently created users in Microsoft Entra (e.g., within the last n days).
5. Exports the report result to CSV.
6. The script can be executed with an MFA enabled account too.
7. It can be executed with Certificate-based Authentication (CBA) too.
8. The script is schedular-friendly.  

For detailed Script execution:  https://o365reports.com/2025/04/15/export-all-entra-users-using-powershell/
============================================================================================
\#>
Param
(
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [int]$RecentlyCreatedUsers,
    [Switch]$GuestUsersOnly,
    [Switch]$EnabledUsersOnly,
    [Switch]$DisabledUsersOnly,
    [Switch]$LicensedUsersOnly,
    [Switch]$UnlicensedUsersOnly,
    [Switch]$UnmanagedUsers
    
)
Function Connect_MgGraph
{
 #Check for module installation
 $Module=Get-Module -Name Microsoft.Graph -ListAvailable
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
  Disconnect-MgGraph | Out-Null
 }


 Write-Host Connecting to Microsoft Graph...
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
 }
 else
 {
  Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"  -NoWelcome
 }
}
Connect_MgGraph


$Location=Get-Location 
$ExportCSV = "$Location\EntraUsers_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$Count=0
$PrintedUsers=0
$RequiredProperties=@('UserPrincipalName','LastPasswordChangeDateTime','AccountEnabled','Country','Department','Jobtitle','SigninActivity','DisplayName','UserType','CreatedDateTime')

Write-Host Generating Entra users report...
Get-MgUser -All -Property $RequiredProperties  | foreach {
 $Print=1
 $UPN=$_.UserPrincipalName
 $DisplayName=$_.DisplayName
 $Count++
 Write-Progress -Activity "`n     Processed users: $Count - $UPN "
 $LastPwdSet=$_.LastPasswordChangeDateTime
 $AccountEnabled=$_.AccountEnabled
 if($AccountEnabled -eq $true)
 {
  $SigninStatus="Allowed"
 }
 else
 {
  $SigninStatus="Denied"
 }

 $SKUs = (Get-MgUserLicenseDetail -UserId $UPN).SkuPartNumber
 $Sku= $SKUs -join ","
 $Department=$_.Department
 $JobTitle=$_.JobTitle
 $LastSigninTime=($_.SignInActivity).LastSignInDateTime
 $LastNonInteractiveSignIn=($_.SignInActivity).LastNonInteractiveSignInDateTime
 $Manager=(Get-MgUserManager -UserId $UPN -ErrorAction SilentlyContinue)
 $ManagerDetails=$Manager.AdditionalProperties
 $ManagerName=$ManagerDetails.userPrincipalName
 $Country= $_.Country
 $CreationTime=$_.CreatedDateTime
 $CreatedSince=(New-TimeSpan -Start $CreationTime).Days
 $UserType=$_.UserType

 #Filter for guest users
 if($GuestUsersOnly.IsPresent -and ($UserType -ne "Guest"))
 { 
  $Print=0
 }
 #Filter for recently created users
 if(($RecentlyCreatedUsers -ne "") -and ($CreatedSince -gt $RecentlyCreatedUsers))
 { 
  $Print=0
 }
 #Filter for sign-in allowed users
 if($EnabledUsersOnly.IsPresent -and ($AccountEnabled -eq $false))
 {
  $Print=0
 }
 #Filter for sign-in disabled users
 if($DisabledUsersOnly.IsPresent -and ($AccountEnabled -eq $true))
 {
  $Print=0
 }
 #Filter for licensed users
 if(($LicensedUsersOnly.IsPresent) -and ($Sku.Length -eq 0))
 {
  $Print=0
 }
 #Filter for unlicensed users
 if(($UnlicensedUsersOnly.IsPresent) -and ($Sku.Length -ne 0))
 {
  $Print=0
 }
 #Filter for users withour manager
 if(($UnmanagedUsers.IsPresent) -and ($Manager -ne $null))
 {
  $Print=0
 }
 
 #Export users based on the given criteria
 if($Print -eq 1)
 {
  $PrintedUsers++
  $Result=[PSCustomObject]@{'Name'=$UPN;'Display Name'=$DisplayName;'User Type'=$UserType;'Sign-in Status'=$SigninStatus;'License'=$SKU;'Department'=$Department;'Job Title'=$JobTitle;'Country'=$Country;'Manager'=$ManagerName;'Pwd Last Change Date'=$LastPwdSet;'Last Signin Date'=$LastSigninTime;'Last Non-interactive Signin Date'=$LastNonInteractiveSignIn;'Creation Time'=$CreationTime}
  $Result | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

Disconnect-MgGraph | Out-Null
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  
#Open Output file after execution
 if((Test-Path -Path $ExportCSV) -eq "True") 
 {
  Write-Host `The exported report contains $PrintedUsers users.
  Write-Host `nEntra users report available in: -NoNewline -Foregroundcolor Yellow; Write-Host $ExportCSV
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
  Write-Host No users found for the given criteria.
 }