<#
=============================================================================================
Name:           Export Microsoft 365 Users Last Successful SignIn Time Report
Version:        1.0
website:        admindroid.com
For detailed Script execution:  https://blog.admindroid.com/get-last-successful-sign-in-date-report-for-microsoft-365-users
============================================================================================
#>
Param
(
    [int]$InactiveDays,
    [switch]$EnabledUsersOnly,
    [switch]$DisabledUsersOnly,
    [switch]$LicensedUsersOnly,
    [switch]$ExternalUsersOnly,
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
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
  Connect-MgGraph -Scopes "User.Read.All","AuditLog.read.All"  -NoWelcome
 }
}
Connect_MgGraph

$ExportCSV = ".\M365User_LastSuccessfulSigninDate_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult=""   
$ExportResults=@()  


#Get friendly name of license plan from external file
$FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData

$Count=0
$PrintedUser=0
#retrieve users
$RequiredProperties=@('UserPrincipalName','EmployeeId','CreatedDateTime','AccountEnabled','Department','JobTitle','SigninActivity')
Get-MgBetaUser -All -Property $RequiredProperties | select $RequiredProperties | ForEach-Object {
 $Count++
 $UPN=$_.UserPrincipalName
 Write-Progress -Activity "`n     Processing user: $Count - $UPN"
 $EmployeeId=$_.EmployeeId
 
 [String]$LastSuccessfulSigninDate=$_.SignInActivity.AdditionalProperties.lastSuccessfulSignInDateTime
 $LastInteractiveSignIn=$_.SignInActivity.LastSignInDateTime
 $LastNon_InteractiveSignIn=$_.SignInActivity.LastNonInteractiveSignInDateTime
 $CreatedDate=$_.CreatedDateTime
 $AccountEnabled=$_.AccountEnabled
 $Department=$_.Department
 $JobTitle=$_.JobTitle
 
 #Calculate Inactive days
 if($LastSuccessfulSigninDate -eq "")
 { 
  $LastSuccessfulSigninDate="Data not available"
  $InactiveUserDays= "-"
 }
 else
 {
 [dateTime]$LastSuccessfulSigninDate=$_.SignInActivity.AdditionalProperties.lastSuccessfulSignInDateTime
  $InactiveUserDays= (New-TimeSpan -Start $LastSuccessfulSigninDate).Days
 }
 
 if($AccountEnabled -eq $true)
 {
  $AccountStatus='Enabled'
 }
 else
 {
  $AccountStatus='Disabled'
 }

 #Get licenses assigned to mailboxes
 $Licenses = (Get-MgBetaUserLicenseDetail -UserId $UPN).SkuPartNumber
 $AssignedLicense = @()

 #Convert license plan to friendly name
 if($Licenses.count -eq 0)
 {
  $LicenseDetails = "No License Assigned"
 }
 else
 {
  foreach($License in $Licenses)
  {
   $EasyName = $FriendlyNameHash[$License]
   if(!($EasyName))
   {$NamePrint = $License}
   else
   {$NamePrint = $EasyName}
   $AssignedLicense += $NamePrint
  }
  $LicenseDetails = $AssignedLicense -join ", "
 }
 $Print=1


 #Inactive days based on successful signins filter
 if($InactiveUserDays -ne "-")
 {
  if(($InactiveDays -ne "") -and ($InactiveDays -gt $InactiveUserDays))
  {
   $Print=0
  }
 }
    



 #Filter for external users
 if(($ExternalUsersOnly.IsPresent) -and ($UPN -notmatch '#EXT#'))
 {
   $Print=0
 }
 
 #Signin Allowed Users
 if($EnabledUsersOnly.IsPresent -and $AccountStatus -eq 'Disabled')
 {      
  $Print=0
 }

 #Signin disabled users
 if($DisabledUsersOnly.IsPresent -and $AccountStatus -eq 'Enabled')
 {
  $Print=0
 }

 #Licensed Users ony
 if($LicensedUsersOnly -and $Licenses.Count -eq 0)
 {
  $Print=0
 }

 #Export users to output file
 if($Print -eq 1)
 {
  $PrintedUser++
  $ExportResult=[PSCustomObject]@{'UPN'=$UPN;'Creation Date'=$CreatedDate;'Last Successful Signin Date'=$LastSuccessfulSigninDate;'Inactive Days'=$InactiveUserDays;'Last Interactive SignIn Date'=$LastInteractiveSignIn;'Last Non Interactive SignIn Date'=$LastNon_InteractiveSignIn;'Emp id'=$EmployeeId;'License Details'=$LicenseDetails;'Account Status'=$AccountStatus;'Department'=$Department;'Job Title'=$JobTitle}
  $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

#Open output file after execution
Write-Host `nScript executed successfully
if((Test-Path -Path $ExportCSV) -eq "True")
{
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    Write-Host "Exported report has $PrintedUser user(s)" 
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6)
    {
        Invoke-Item "$ExportCSV"
    }
    Write-Host "Detailed report available in: $ExportCSV"
}
else
{
    Write-Host "No user found" -ForegroundColor Red
}