<#
=============================================================================================
Name:           Get all enterprise apps and their owners 
Version:        1.1
Website:        o365reports.com

-------------------
Script Highlights: 
-------------------
1. The script exports all enterprise apps along with its owners in Microsoft Entra.  
2. Generates report for sign-in enabled applications alone. 
3. Exports report for sign-in disabled applications only. 
4. Filters applications that are hidden from all users except assigned users. 
5. Provides the list of applications that are visible to all users in the organization. 
6. Lists applications that are accessible to all users in the organization.  
7. Identifies applications that can be accessed only by assigned users. 
8. Fetches the list of ownerless applications in Microsoft Entra. 
9. Assists in filtering home tenant applications only. 
10. Exports applications from external tenants only. 
11. The script uses MS Graph PowerShell and installs MS Graph PowerShell SDK (if not installed already) upon your confirmation.  
12. Exports the report result to CSV. 
13. The script can be executed with an MFA enabled account too. 
14. It can be executed with certificate-based authentication (CBA) too. 
15. The script is schedular-friendly.

------------
Change Log:
------------
V1.0 (Nov 26, 2024) - File created.
V1.1 (Apr 08, 2026) - Error handling case added for app creation time.

For detailed Script execution: https://o365reports.com/2024/11/26/export-all-enterprise-apps-and-their-owners-in-microsoft-entra/

============================================================================================
#>
Param
(
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$SigninEnabledAppsOnly,
    [Switch]$SigninDisabledAppsOnly,
    [Switch]$HiddenApps,
    [Switch]$VisibleToAllUsers,
    [Switch]$AccessScopeToAllUsers,
    [Switch]$RoleAssignmentRequiredApps,
    [Switch]$OwnerlessApps,
    [Switch]$HomeTenantAppsOnly,
    [Switch]$ExternalTenantAppsOnly
)
Function Connect_MgGraph
{
 #Check for module installation
 $Module=Get-Module -Name Microsoft.Graph -ListAvailable
 if($Module.count -eq 0){ 
  Write-Host "Microsoft Graph PowerShell SDK is not available"  -ForegroundColor yellow  
  $Confirm= Read-Host "Are you sure you want to install module? [Y] Yes [N] No"
  if($Confirm -match "^[yY]"){ 
   Write-host "Installing Microsoft Graph PowerShell module..."
   Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
  }
  else{
   Write-Host "Microsoft Graph PowerShell module is required to run this script. Please install module using 'Install-Module Microsoft.Graph' cmdlet." 
   Exit
  }
 }
 #Disconnect Existing MgGraph session
 if($CreateSession.IsPresent){
  Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
 }


 Write-Host "`nConnecting to Microsoft Graph..."
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne "")){  
  Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
 }
 else{
  Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome
 }
}
Connect_MgGraph

$Location = Get-Location
$ExportCSV = "$Location\EnterpriseApps_and_their_Owners_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$PrintedCount = 0
$Count = 0
$TenantGUID = (Get-MgOrganization).Id

Get-MgServicePrincipal -All | foreach{
 $Print = 1
 $Count++
 $EnterpriseAppName = $_.DisplayName
 Write-Progress -Activity "`nProcessed enterprise apps: $($Count) - $($EnterpriseAppName)"
 $UserSigninStatus = $_.AccountEnabled
 $Id = $_.Id
 $Tags = $_.Tags
 if($Tags -contains "HideApp"){
  $UserVisibility = "Hidden"
 }
 else{
  $UserVisibility="Visible"
 }
 $IsRoleAssignmentRequired=$_.AppRoleAssignmentRequired
 if($IsRoleAssignmentRequired -eq $true){
  $AccessScope="Only assigned users can access"
 }
 else{
  $AccessScope="All users can access"
 }
 $CreationTime = $null
 $CreationTime = $_.AdditionalProperties.createdDateTime -as [datetime]
 if($CreationTime){ $CreationTime = $CreationTime.ToLocalTime() }
 $ServicePrincipalType=$_.ServicePrincipalType
 $AppRegistrationName=$_.AppDisplayName
 $AppOwnerOrgId=$_.AppOwnerOrganizationId
 if($AppOwnerOrgId -eq $TenantGUID){
  $AppOrigin="Home tenant"
 }
 else{
  $AppOrigin="External tenant"
 }
 $Owners = ((Get-MgServicePrincipalOwner -ServicePrincipalId $Id -ErrorAction SilentlyContinue).AdditionalProperties.userPrincipalName | Where-Object { $_ }) -join ","
 if (-not $Owners) { $Owners = "-" }

 #Filtering the result
 if(($SigninEnabledAppsOnly.IsPresent) -and ($UserSigninStatus -eq $false)){
  $Print=0
 }
 elseif(($SigninDisabledAppsOnly.IsPresent) -and ($UserSigninStatus -eq $true)){
  $Print=0
 }
 if(($HiddenApps.IsPresent) -and ($UserVisibility -eq "Visible")){
  $Print=0
 }
 elseif(($VisibleToAllUsers.IsPresent) -and ($UserVisibility -eq "Hidden")){
  $Print=0
 }
 if(($AccessScopeToAllUsers.IsPresent) -and ($AccessScope -eq "Only assigned users can access")){
  $Print=0
 }
 elseif(($RoleAssignmentRequiredApps.IsPresent) -and ($AccessScope -eq "All users can access")){
  $Print=0
 }
 if(($OwnerlessApps.IsPresent) -and ($Owners -ne "-")){
  $Print=0
 }
 if(($HomeTenantAppsOnly.IsPresent) -and ($AppOrigin -eq "External tenant")){
  $Print=0
 }
 elseif(($ExternalTenantAppsOnly.IsPresent) -and ($AppOrigin -eq "Home tenant")){
  $Print=0
 }

 if($Print -eq 1){
    $PrintedCount++
    $ExportResult = [PSCustomObject]@{
    'Enterprise App Name'      = $EnterpriseAppName
    'App Id'                   = $Id
    'App Owners'               = $Owners
    'App Creation Time'        = $CreationTime
    'User Signin Allowed'      = $UserSigninStatus
    'User Visibility'          = $UserVisibility
    'Role Assignment Required' = $AccessScope
    'Service Principal Type'   = $ServicePrincipalType
    'App Registration Name'    = $AppRegistrationName
    'App Origin'               = $AppOrigin
    'App Org Id'               = $AppOwnerOrgId
    }
    $ExportResult | Export-Csv -Path $ExportCSV -NoTypeInformation -Append
  }
}

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Write-Host "`n~~ Script prepared by AdminDroid Community ~~`n" -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green
 
#Open output file after execution 
 If($PrintedCount -eq 0){
  Write-Host "No data found for the given criteria."
 }
 else{
  Write-Host "`nThe script processed $($Count) enterprise apps and the output file contains $($PrintedCount) records."
  if((Test-Path -Path $ExportCSV) -eq "True"){
   Write-Host "`nThe Output file available in: " -NoNewline -ForegroundColor Yellow
   Write-Host $ExportCSV 
   $Prompt = New-Object -ComObject wscript.shell      
   $UserInput = $Prompt.popup("Do you want to open output file?",0,"Open Output File",4)   
   If ($UserInput -eq 6){   
    Invoke-Item "$ExportCSV"   
   } 
  }
 }