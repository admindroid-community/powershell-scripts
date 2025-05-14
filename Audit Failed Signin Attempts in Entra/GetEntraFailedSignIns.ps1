<#
=============================================================================================
Name:           Export Microsoft 365 Sign-in Failure Report Using PowerShell 
Version:        1.0
Website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~

1. The script automatically verifies and installs the Microsoft Graph PowerShell SDK module (if not installed already) upon your confirmation. 
2. Generates a report that retrieves all failed interactive sign-in attempts by default.
3. Enables filtering of failed attempts from the following sign-in types.  
    -> Non-interactive user sign-ins 
    -> Service principal sign-ins 
    -> Managed identity sign-ins
    -> All type of sign-in failure events 
4. Allows you to export a failed login attempts report for a custom period. 
5. View sign-in failure events based on error code.  
6. Helps export failed risky sign-ins alone. 
7. Exports a report on guest users’ sign-in failures.
8. Segments failed sign-in attempts based on MFA enforced & not enforced. 
9. You can export the report to choose either All users or Specific user(s). 
10. The script can be executed with an MFA-enabled account too. 
11. Exports report results as a CSV file.
12. The script is scheduler friendly.
13. It can be executed with certificate-based authentication (CBA) too.


For detailed script execution: https://o365reports.com/2025/05/13/export-microsoft-365-sign-in-failure-report-using-powershell/
 
============================================================================================
#>

Param
(
    [nullable[datetime]]$StartDate, #Pass value as YYYY-MM-DD
    [nullable[datetime]]$EndDate, #Pass value as YYYY-MM-DD
    [nullable[int]]$ErrorCode=$null,
    [ValidateSet( "NonInteractiveSignIns", "ServicePrincipalSignIns", "ManagedIdentitySignIns", "AllSignIns" )]
    [string]$SignInEventType,
    [switch]$RiskySignInsOnly,
    [switch]$GuestUserSignInsOnly,
    [switch]$MFASignInsOnly,
    [switch]$NonMFASignInsOnly,
    [switch]$CreateSession,
    [string[]]$UserPrincipalName,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Function Connect_MgGraph
{
 #Check for module installation
 $Module=Get-Module -Name microsoft.graph.beta.reports -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Microsoft Graph Beta PowerShell SDK is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Microsoft Graph Beta PowerShell module..."
   Install-Module Microsoft.Graph.beta.reports -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
  }
  else
  {
   Write-Host "Microsoft Graph Beta PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph.beta.reports cmdlet." 
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
  Connect-MgGraph -Scopes "AuditLog.Read.All", "Policy.Read.ConditionalAccess" -NoWelcome
 }
}
Connect_MgGraph

$Location = Get-Location
$CurrentDate = Get-Date
$ExportCSV = "$Location\M365Users_Failed_SignIn_Report$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult =""   
$ExportResults = @() 
$SignInFilter = $Filter = $null

if ($StartDate -ne $null) {
    $ThresholdDate = ($CurrentDate).AddDays(-30)
    if ($StartDate -lt $ThresholdDate) {
        Write-Host "Error: The StartDate cannot be earlier than $($ThresholdDate.ToString('yyyy-MM-dd'))." -ForegroundColor Red
        Exit
    } else {
        $SignInFilter += " and (createdDateTime ge $($StartDate.ToString('yyyy-MM-dd')))"
    }
} 

if ($EndDate -ne $null) {
    $SignInFilter += " and (createdDateTime le $($EndDate.AddDays(1).ToString("yyyy-MM-dd")))"
} 

if ($GuestUserSignInsOnly.IsPresent) {
    $SignInFilter += " and (UserType eq 'guest')"
} 

switch ($SignInEventType) {
    #"InteractiveSignIns" { $SignInFilter += " and (signInEventTypes/any(t: t eq 'interactiveUser'))" }
    "NonInteractiveSignIns" { $SignInFilter += " and (signInEventTypes/any(t: t eq 'noninteractiveUser'))" }
    "ServicePrincipalSignIns" { $SignInFilter += " and (signInEventTypes/any(t: t eq 'servicePrincipal'))" }
    "ManagedIdentitySignIns" { $SignInFilter += " and (signInEventTypes/any(t: t eq 'managedIdentity'))" }
    "AllSignIns" { $SignInFilter += " and ((signInEventTypes/any(t: t eq 'interactiveUser')) or (signInEventTypes/any(t: t eq 'nonInteractiveUser')) or (signInEventTypes/any(t: t eq 'servicePrincipal')) or (signInEventTypes/any(t: t eq 'managedIdentity')))" }
} 

if ($NonMFASignInsOnly.IsPresent) {
    $SignInFilter += " and (authenticationRequirement eq 'singleFactorAuthentication')"
}
if ($MFASignInsOnly.IsPresent) {
    $SignInFilter += " and (authenticationRequirement eq 'multiFactorAuthentication')"    
}

if (!([string]::IsNullOrEmpty($ErrorCode))) { 
    $Filter = "(Status/Errorcode eq $($ErrorCode)) $($SignInFilter)"
} else {
    $Filter = "(Status/Errorcode ne 0) $($SignInFilter)"
}

$Count=0
$PrintedLogs=0
#retrieve signin activities
Write-Host "Generating M365 users' failed signin report..."
Get-MgBetaAuditLogSignIn -All -Filter "$($Filter)" | ForEach-Object {
 $Count++
 $UPN=$_.UserPrincipalName
 Write-Progress -Activity "`n     Processed sign-in record: $count"
 $CreatedDate=$_.CreatedDateTime
 $Id=$_.Id
 $UserDisplayName=$_.UserDisplayName
 $UPN=$_.UserPrincipalName
 $Authentication=$_.AuthenticationRequirement
 $Location="$($_.Location.City), $($_.Location.State), $($_.Location.CountryOrRegion)"
 $DeviceName=$_.DeviceDetail.DisplayName
 $Browser=$_.DeviceDetail.Browser
 $OperatingSystem=$_.DeviceDetail.OperatingSystem
 $IpAddress=$_.IpAddress
 $FlaggedForReview = $_.FlaggedForReview
 $SignInErrorCode=$_.Status.ErrorCode
 $FailureReason=$_.Status.FailureReason
 $AdditionalDetails = $_.Status.AdditionalDetails
 $UserType=$_.UserType
 $IsInteractive=$_.IsInteractive
 $RiskDetail=$_.RiskDetail
 $RiskState=$_.RiskState
 $AppDispalyName=$_.AppDisplayName
 $ResourceDisplayName=$_.ResourceDisplayName
 $ConditionalAccessPolicyStatus=$_.ConditionalAccessStatus
 $AppliedConditionalAccessPolicies=$_.AppliedConditionalAccessPolicies
 

 $AppliedPolicies = @()
 $AppliedConditionalAccessPolicies | ForEach-Object {
    if($_.Result -eq 'success' -or $_.Result -eq 'failure'){
        $AppliedPolicies += $_.DisplayName
    }
 }

 if($AppliedPolicies.Count -eq 0){
 $AppliedPolicies = "None"
 } else{
 $AppliedPolicies = $AppliedPolicies -join ", "
 }

 if($SignInErrorCode -ne 0) { $Status='Failed' }

 if($FailureReason -eq 'Other.') {
   $FailureReason="Others"
 }

 #Declaring Print flag for all sign-ins 
 $Print=1

 #Filter for risky sign-ins
 if(($RiskySignInsOnly.IsPresent) -and ($RiskDetail -eq 'none'))
 {
   $Print=0
 }

 #Filter for specific user sign-ins
 if(($UserPrincipalName -ne $null) -and ($UserPrincipalName -notcontains $UPN))
 {
   $Print=0
 }


  #Export users to output file
 if($Print -eq 1)
 {
  $PrintedLogs++
  $ExportResult=[PSCustomObject]@{'Signin Date'=$CreatedDate; 'UPN'=$UPN; 'User Name'=$UserDisplayName; 'Application Name'=$AppDispalyName; 'Resource Name'=$ResourceDisplayName; 'Status'=$Status; 'Error Code'=$SignInErrorCode; 'Failure Reason'=$FailureReason; 'User type'=$UserType; 'Ip Address'=$IpAddress; 'Location'=$Location; 'Device Name'=$DeviceName; 'Browser'=$Browser; 'Operating System'=$OperatingSystem; 'Authentication Requirement'=$Authentication; 'Conditional access status'=$ConditionalAccessPolicyStatus;  'Applied Conditional Access Policies'=$AppliedPolicies; 'Risk detail'=$RiskDetail; 'Risk state'=$RiskState; 'SigninId'=$Id; 'Flagged For Review'=$FlaggedForReview; 'IsInteractive'=$IsInteractive; 'AdditionalDetails'=$AdditionalDetails;}
  $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

#Disconnect the session after execution
Disconnect-MgGraph | Out-Null

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n

#Open output file after execution
if((Test-Path -Path $ExportCSV) -eq "True")
{   
    Write-Host " The Output file availble in: " -NoNewline -ForegroundColor Yellow; Write-Host "$ExportCSV" `n 
    Write-Host " Exported report has $PrintedLogs failed signin activities." 
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6)
    {
        Invoke-Item "$ExportCSV"
    }
}
else
{
    Write-Host "No logs found" 
}