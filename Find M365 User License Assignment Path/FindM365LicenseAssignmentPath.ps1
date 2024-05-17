<#
=============================================================================================
Name:           Find & Export Microsoft 365 User License Assignment Paths using PowerShell 
Version:        1.0
website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script uses MS Graph PowerShell and installs MS Graph PowerShell SDK (if not installed already) upon your confirmation. 
2. The script can be executed with MFA enabled account too. 
3. Exports directly assigned licenses alone. 
4. Exports group-based license assignments alone. 
5. Helps to identify users with license assignment errors. 
6. Converts SKU name into user-friendly name. 
7. Produces a list of disabled service plans for the assigned license. 
8. Exports report results as a CSV file. 
9. The script is scheduler friendly. 
10. It can be executed with certificate-based authentication (CBA) too. 


For detailed Script execution:  https://o365reports.com/2024/05/14/find-export-microsoft-365-user-license-assignment-paths-using-powershell/
============================================================================================
#>
Param
(
    [switch]$ShowDirectlyAssignedLicenses,
    [switch]$ShowGrpBasedLicenses,
    [switch]$DisabledUsersOnly,
    [Switch]$FindUsersWithLicenseAssignmentErrors,
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

Function Convert-FrndlyName {
    param(
        [Parameter(Mandatory=$true)]
        [Array]$InputIds
    )
    $EasyName = $FriendlyNameHash[$SkuName]
   if(!($EasyName))
   {$NamePrint = $SkuName}
   else
   {$NamePrint = $EasyName}
   return $NamePrint
}
Connect_MgGraph
$Location=Get-Location
$ExportCSV="$Location\M365Users_LicenseAssignmentPath_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult=""   
$ExportResults=@() 
$PrintedUser=0 

#Get license in the organization and saving it as hash table
$SKUHashtable=@{}
Get-MgBetaSubscribedSku –All | foreach{
 $SKUHashtable[$_.skuid]=$_.Skupartnumber
}
#Get friendly name of Subscription plan from external file
$FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData
#Get friendly name of Service plan from external file
$ServicePlanHash=@{}
Import-Csv -Path .\ServicePlansFrndlyName.csv | ForEach-Object {
 $ServicePlanHash[$_.ServicePlanId] = $_.ServicePlanFriendlyNames
}

$GroupNameHash=@{}
#Process users
$RequiredProperties=@('UserPrincipalName','DisplayName','EmployeeId','CreatedDateTime','AccountEnabled','Department','JobTitle','LicenseAssignmentStates','AssignedLicenses','SigninActivity')
Get-MgBetaUser -All -Property $RequiredProperties | select $RequiredProperties | ForEach-Object {
 $Count++
 $Print=1
 $DirectlyAssignedLicense=@()
 $GroupBasedLicense=@()
 $DirectlyAssignedLicense_FrndlyName=@()
 $GroupBasedLicense_FrndlyName=@()
 $UPN=$_.UserPrincipalName
 Write-Progress -Activity "`n     Processing user: $Count - $UPN"
 $DisplayName=$_.DisplayName
 $AccountEnabled=$_.AccountEnabled
 $Department=$_.Department
 $JobTitle=$_.JobTitle
 $LastSignIn=$_.SignInActivity.LastSignInDateTime
 if($LastSignIn -eq $null)
 {
  $LastSignIn = "Never Logged In"
  $InactiveDays = "-"
 }
 else
 {
  $InactiveDays = (New-TimeSpan -Start $LastSignIn).Days
 }
 $LicenseAssignmentStates=$_.LicenseAssignmentStates

 if($AccountEnabled -eq $true)
 {
  $AccountStatus='Enabled'
 }
 else
 {
  $AccountStatus='Disabled'
 }

 foreach($License in $licenseAssignmentStates)
 { 
  $SkuName=$SkuHashtable[$License.SkuId]
  $FriendlyName=Convert-FrndlyName -InputIds $SkuName
  $DisabledPlans=$License.DisabledPlans
  $ServicePlanNames=@()
  if($DisabledPlans.count -ne 0 )
  {
   foreach($DisabledPlan in $DisabledPlans)
   {
    $ServicePlanName = $ServicePlanHash[$DisabledPlan]
    if(!($ServicePlanName))
    {$NamePrint = $DisabledPlan}
    else
    {$NamePrint = $ServicePlanName}
    $ServicePlanNames += $NamePrint
   }
  }
  $DisabledPlans=$ServicePlanNames -join ","
  $State=$License.State
  $Error=$License.Error
 
  #Filter for users with license assignment errors
  if($FindUsersWithLicenseAssignmentErrors.IsPresent -and ($State -eq "Active"))
  {
   $Print=0
  }

  if($License.AssignedByGroup -eq $null)
  {
   $LicenseAssignmentPath="Directly assigned"
   $GroupName="NA"
   #Filter for group based license assignment
   if($ShowGrpBasedLicenses.IsPresent)
   {
    $Print=0
   }
  }
  else
  {
   $LicenseAssignmentPath="Inherited from group"

   #Filter for directly assigned licenses
   if($ShowDirectlyAssignedLicenses.IsPresent)
   {
    $Print=0
   }

   $AssignedByGroup=$License.AssignedByGroup
   # Check Id-Name pair already exist in hash table
   if($GroupNameHash.ContainsKey($AssignedByGroup))
   {
    $GroupName=$GroupNameHash[$AssignedByGroup]
   }
   else
   {
    $GroupName=(Get-MgBetagroup -GroupId $AssignedByGroup).DisplayName
    $GroupNameHash[$AssignedByGroup]=$GroupName
   }
  }
  if($Print -eq 1)
  {
   $ExportResult=[PSCustomObject]@{'Display Name'=$DisplayName;'UPN'=$UPN;'License Assignment Path'=$LicenseAssignmentPath;'Sku Name'=$SkuName;'Sku_FriendlyName'=$FriendlyName;'Disabled Plans'=$DisabledPlans;'Assigned via(group name)'=$GroupName;'State'=$State;'Error'=$Error;'Last Signin Time'=$LastSignIn;'Inactive Days'=$InactiveDays;'Account Status'=$AccountStatus;'Department'=$Department;'Job Title'=$JobTitle}
   $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
  }
 }
}
#Open output file after execution
Write-Host `nScript executed successfully
if((Test-Path -Path $ExportCSV) -eq "True")
{
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6)
    {
        Invoke-Item "$ExportCSV"
    }
    Write-Host "Detailed report available in: $ExportCSV" -ForegroundColor Cyan
}
else
{
    Write-Host "No user found" -ForegroundColor Red
}



































