<#
=============================================================================================
Name:           Export Microsoft 365 Users’ Sign-in Report Using PowerShell 
Version:        1.0
Website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Exports all Entra ID sign-in logs in user-friendly format. 
2. Allows you to findout successful and failed sign-in attempts separately. 
3. Filters interactive/non-interactive user sign-ins. 
4. Helps export risky sign-ins alone. 
5. Tracks guest users’ sign-in history. 
6. Segments sign-in attempts based on Conditional Access applied & not applied. 
7. Helps to monitor CA policy sign-in failures & success separately. 
8. You can export the report to choose either ‘All users’ or ‘Specific user(s)’. 
9. The script uses MS Graph PowerShell and installs MS Graph Beta PowerShell SDK (if not installed already) upon your confirmation.  
10. The script can be executed with an MFA-enabled account too.  
11. Exports report results as a CSV file.  
12. The script is scheduler friendly.  
13. It can be executed with certificate-based authentication (CBA) too. 

For detailed script execution:https://o365reports.com/2024/07/02/export-microsoft-365-users-sign-in-report-using-powershell/
 
============================================================================================
#>
Param
(
    [switch]$RiskySignInsOnly,
    [switch]$GuestUserSignInsOnly,
    [switch]$Success,
    [switch]$Failure,
    [switch]$InteractiveOnly,
    [switch]$NonInteractiveOnly,
    [switch]$CAPNotAppliedOnly,
    [switch]$CAPAppliedOnly,
    [switch]$CAPSuccessOnly,
    [switch]$CAPFailedOnly,
    [switch]$CreateSession,
    [string[]]$UserPrincipalName,
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
  Connect-MgGraph -Scopes "AuditLog.Read.All", "Directory.Read.All", "Policy.Read.ConditionalAccess"  -NoWelcome
 }
}
Connect_MgGraph

$Location=Get-Location
$ExportCSV = "$Location\M365Users_Signin_Report$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult=""   
$ExportResults=@()  


$Count=0
$PrintedLogs=0
#retrieve signin activities
Write-Host "Generating M365 users' signin report..."
Get-MgBetaAuditLogSignIn -All | ForEach-Object {
 $Count++
 $UPN=$_.UserPrincipalName
 Write-Progress -Activity "`n     Processed sign-in record: $count "
 $CreatedDate=$_.CreatedDateTime
 $Id=$_.Id
 $UserDisplayName=$_.UserDisplayName
 $UPN=$_.UserPrincipalName
 $AuthenticationRequirement=$_.AuthenticationRequirement
 $Location="$($_.Location.City),$($_.Location.State),$($_.Location.CountryOrRegion)"
 $DeviceName=$_.DeviceDetail.DisplayName
 $Browser=$_.DeviceDetail.Browser
 $OperatingSystem=$_.DeviceDetail.OperatingSystem
 $IpAddress=$_.IpAddress
 $ErrorCode=$_.Status.ErrorCode
 $FailureReason=$_.Status.FailureReason
 $UserType=$_.UserType
 $RiskDetail=$_.RiskDetail
 $IsInteractive=$_.IsInteractive
 $RiskState=$_.RiskState
 $AppDispalyName=$_.AppDisplayName
 $ResourceDisplayName=$_.ResourceDisplayName
 $ConditionalAccessStatus=$_.ConditionalAccessStatus
 $AppliedConditionalAccessPolicies=$_.AppliedConditionalAccessPolicies
 


 $AppliedPolicies = @()
 $AppliedConditionalAccessPolicies | ForEach-Object {
    if($_.Result -eq 'Success' -or $_.Result -eq 'Failed'){
        $AppliedPolicies += $_.DisplayName
    }
 }
 if($AppliedPolicies.Count -eq 0){
 $AppliedPolicies = "None"
 } else{
 $AppliedPolicies = $AppliedPolicies -join ", "
 }

 
 if($ErrorCode -eq 0)
 {
  $Status='Success'
 }
 else
 {
  $Status='Failed'
 }

 if($FailureReason -eq 'Other.') {
   $FailureReason="None"
 }

 #Declaring Print flag for all sign-ins 
 $Print=1

 #Filter for successful sign-ins
 if(($Success.IsPresent) -and ($Status -ne 'Success'))
 {
   $Print=0
 }

 #Filter for failed sign-ins
 if(($Failure.IsPresent) -and ($Status -ne 'Failed'))
 {
   $Print=0
 }

 #Filter for applied conditional access policies
 if(($CAPAppliedOnly.IsPresent) -and ($ConditionalAccessStatus -eq 'NotApplied'))
 {
   $Print=0
 }

 #Filter for not applied conditional access policies
 if(($CAPNotAppliedOnly.IsPresent) -and ($ConditionalAccessStatus -ne 'NotApplied'))
 {
   $Print=0
 }

 #Filter for failed conditional access policies
 if(($CAPFailedOnly.IsPresent) -and ($ConditionalAccessStatus -ne 'Failed'))
 {
   $Print=0
 }

 #Filter for succeeded conditional access policies
 if(($CAPSuccessOnly.IsPresent) -and ($ConditionalAccessStatus -ne 'Success'))
 {
   $Print=0
 }

 #Filter for risky sign-ins
 if(($RiskySignInsOnly.IsPresent) -and ($RiskDetail -eq 'none'))
 {
   $Print=0
 }

 #Filter for guest user sign-ins
 if(($GuestUserSignInsOnly.IsPresent) -and ($UserType -eq 'member'))
 {
   $Print=0
 }  

 #Filter for guest user sign-ins
 if(($UserPrincipalName -ne $null) -and ($UserPrincipalName -notcontains $UPN))
 {
   $Print=0
 }

 
 #Filter for interactive sign-ins
 if(($InteractiveOnly.IsPresent) -and (!$IsInteractive))
 {
   $Print=0
 }
 

 #Filter for non-interactive sign-ins
 if(($NonInteractiveOnly.IsPresent) -and ($IsInteractive))
 {
   $Print=0
 }


  #Export users to output file
 if($Print -eq 1)
 {
  $PrintedLogs++
  $ExportResult=[PSCustomObject]@{'Signin Date'=$CreatedDate; 'User Name'=$UserDisplayName; 'SigninId'=$Id;'UPN'=$UPN; 'Status'=$Status; 'Ip Address'=$IpAddress; 'Location'=$Location; 'Device Name'=$DeviceName; 'Browser'=$Browser; 'Operating System'=$OperatingSystem; 'User type'=$UserType; 'Authentication Requirement'=$AuthenticationRequirement; 'Risk detail'=$RiskDetail; 'Risk state'=$RiskState; 'Conditional access status'=$ConditionalAccessStatus;  'Applied Conditional Access Policies'=$AppliedPolicies; 'IsInteractive'=$IsInteractive;}
  $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

#Disconnect the session after execution
Disconnect-MgGraph | Out-Null

#Open output file after execution

if((Test-Path -Path $ExportCSV) -eq "True")
{   
    Write-Host `n "The Output file availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$ExportCSV" `n 
    Write-Host " Exported report has $PrintedLogs signin activities" 
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
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