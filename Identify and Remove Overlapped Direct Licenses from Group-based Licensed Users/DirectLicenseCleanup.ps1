<#
=============================================================================================
Name:           Remove Direct Licenses for Group-Licensed Users Using PowerShell  
Version:        1.0
website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~ 

1. This script generates reports on users with overlapping direct & group-based license assignments. 
2. The script removes direct assigned license(s) if the same license inherited via groups too.  
3. This script installs MS Graph PowerShell SDK (if not installed already) upon your confirmation. 
4. The script can be executed with an MFA-enabled account too. 
5. The script is schedular-friendly. 
6. It can be executed with certificate-based authentication (CBA) too. 

For detailed Script execution: https://o365reports.com/2024/08/27/remove-direct-licenses-for-group-licensed-users-using-powershell/
============================================================================================
#>
Param
(
    [switch]$GenerateReportOnly,
    [switch]$CreateSession,
    [switch]$Force,
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
$ExportCSV = "$Location\M365Users_Overlapping_License_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$LogFile="$Location\LicenseRemoval_LogFile_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult=""   
$ExportResults=@() 
$PrintedUser=0 
$Count=0

#Get license in the organization and saving it as hash table
$SKUHashtable=@{}
Get-MgBetaSubscribedSku –All | foreach{
$SKUHashtable[$_.skuid]=$_.Skupartnumber}

#Get friendly name of license plan from external file
$FriendlyNameHash=Get-Content -Raw -Path "$Location\LicenseFriendlyName.txt" -ErrorAction Stop | ConvertFrom-StringData

#Confiramtion for removing duplicate license assignment
$LicenseRemovalConfirmation=$false
if($Force.IsPresent)
{
 $LicenseRemovalConfirmation=$true
}
elseif(!($GenerateReportOnly.IsPresent))
{ 
 Write-Host  "Do you want to remove the direct license(s) that inherited via groups? [Y] Yes [N] No " -ForegroundColor Yellow
 $Confirm= Read-Host
 if($Confirm -match "[Y]") 
 { 
  $LicenseRemovalConfirmation=$true
 }
 else
 {
  Write-Host "Proceeding report generation without removing license(s) from users" -ForegroundColor Cyan
 }
}

#Process users
$RequiredProperties=@('UserPrincipalName','DisplayName','EmployeeId','CreatedDateTime','AccountEnabled','Department','JobTitle','LicenseAssignmentStates')
Get-MgBetaUser -Filter "assignedLicenses/`$count ne 0" -All -Property $RequiredProperties -ConsistencyLevel eventual -CountVariable Records | select $RequiredProperties | ForEach-Object {
 $Count++
 $Print=1
 $DirectlyAssignedLicense=@()
 $GroupBasedLicense=@()
 $DirectlyAssignedLicense_FrndlyName=@()
 $GroupBasedLicense_FrndlyName=@()
 $DirectlyAssignedSKUs=@()
 $InheritedSKUs=@()
 $DuplicateAssignment=@()
 $UPN=$_.UserPrincipalName
 Write-Progress -Activity "`n     Processed user: $Count - $UPN"
 $DisplayName=$_.DisplayName
 $AccountEnabled=$_.AccountEnabled
 $Department=$_.Department
 $JobTitle=$_.JobTitle
 $LicenseAssignmentStates=$_.LicenseAssignmentStates


 if($AccountEnabled -eq $true)
 {
  $AccountStatus='Enabled'
 }
 else
 {
  $AccountStatus='Disabled'
 }
foreach($License in $LicenseAssignmentStates)
{
 $SKU=$License.SkuId
 $SkuName=$SkuHashtable[$SKU]
 $FriendlyName=Convert-FrndlyName -InputIds $SkuName
 if($License.AssignedByGroup -eq $null)
 {
  $DirectlyAssignedLicense += $SkuHashtable[$SKU]
  $DirectlyAssignedLicense_FrndlyName += $FriendlyName
  $DirectlyAssignedSKUs +=$SKU
   
 }
 elseif($License.AssignedByGroup -ne $null -and $License.State -eq "Active")
 {
  $GroupBasedLicense += $SkuHashtable[$SKU]
  $GroupBasedLicense_FrndlyName +=$FriendlyName
  $InheritedSKUs +=$SKU
 }
}
 #Check for duplicate license assignment
 if($DirectlyAssignedLicense.Count -ne 0 -and $GroupBasedLicense.Count -ne 0)
 {
  $IsDuplicateLicenseFound = Compare-Object -ReferenceObject $DirectlyAssignedSKUs -DifferenceObject $InheritedSKUs -IncludeEqual -ExcludeDifferent
  if($IsDuplicateLicenseFound.inputObject -ne $null)
  {
   $DuplicateLicenses=$IsDuplicateLicenseFound.inputObject
   $DuplicateLicenseCount=$DuplicateLicenses.count
   foreach ($DuplicateLicense in $DuplicateLicenses)
   {
    $DuplicateAssignment += $SkuHashtable[$DuplicateLicense]
   }
   
   $Print=0
  }
 }

 if($Print -ne 0)
 {
  return
 }  

  
 $DirectlyAssignedLicense=$DirectlyAssignedLicense -join ","
 $GroupBasedLicense=$GroupBasedLicense -join ","
 $DirectlyAssignedLicense_FrndlyName=$DirectlyAssignedLicense_FrndlyName -join ","
 $GroupBasedLicense_FrndlyName=$GroupBasedLicense_FrndlyName -join ","
 $DuplicateAssignment=$DuplicateAssignment -join ","

 #Export output to CSV file
 $PrintedUser++
 $ExportResult=[PSCustomObject]@{'Display Name'=$DisplayName;'UPN'=$UPN;'Directly Assigned Licenses'=$DirectlyAssignedLicense;'Group Based Licenses'=$GroupBasedLicense;'Duplicate License(s)'=$DuplicateAssignment;'Duplicate License Count'=$DuplicateLicenseCount;'Directly Assigned Licenses(Frndly_Name)'=$DirectlyAssignedLicense_FrndlyName;'Group based Licenses(Frndly_Name)'=$GroupBasedLicense_FrndlyName;'Account Status'=$AccountStatus;'Department'=$Department;'Job Title'=$JobTitle}
 $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append

 


 #Remove Direct License Assignement if the same license assigned via group too
 if(($LicenseRemovalConfirmation -eq $true))
 {
  foreach($DupLicense in $DuplicateLicenses)
  { 
   $FriendlyName=Convert-FrndlyName -InputIds $DupLicense
   Write-Progress "Removing $FriendlyName license from the user $UPN"
   Set-MgBetaUserLicense –UserId $UPN -RemoveLicenses @($DupLicense) -AddLicenses @() | Out-Null
   If($?)  
   {  
    "Removing $FriendlyName license from the user $UPN successful" | Out-File $LogFile -Append
   }  
   Else  
   {  
    "Removing $FriendlyName license from the user $UPN failed" | Out-File $LogFile -Append
   }   
  }
 }
}


#Open output file after execution
Write-Host `nScript execution completed
if((Test-Path -Path $ExportCSV) -eq "True")
{
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    Write-Host "Exported report has $PrintedUser user(s)" 
    
    Write-Host "Detailed report available in:" -NoNewline -ForegroundColor Yellow; Write-Host $ExportCSV
    if($LicenseRemovalConfirmation -eq $true)
    {
     Write-Host "License removal log file available in:" -NoNewline -ForegroundColor Yellow; Write-Host $LogFile
    }
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6)
    {
        Invoke-Item "$ExportCSV"
    }

}
else
{
    Write-Host "No user found" -ForegroundColor Red
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
   
}