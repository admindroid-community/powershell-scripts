<#
=============================================================================================
Name         : Find Licensed Groups in Microsoft 365 Using PowerShell  
Version      : 1.0
website      : o365reports.com

-----------------
Script Highlights
-----------------
1. This script allows you to export groups with Microsoft 365 licenses in the organization. 
2. Helps to identify the friendly name of licenses assigned to a group. 
3. The script uses MS Graph PowerShell and installs MS Graph PowerShell SDK (if not installed already) upon your confirmation. 
4. Exports the report result to CSV.  
5. The script can be executed with an MFA enabled account too. 
6. It can be executed with certificate-based authentication (CBA) too. 
7. The script is schedular-friendly. 

For detailed Script execution: https://o365reports.com/2024/09/17/find-licensed-groups-in-microsoft-365-using-powershell/
============================================================================================
#>

Param
(
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
  Connect-MgGraph -Scopes "Directory.Read.All"  -NoWelcome
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
$ExportCSV = "$Location\GroupsThatAutoAssignLicenses_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
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
#Get friendly name of Service plan from external file
$ServicePlanHash=@{}
Import-Csv -Path .\ServicePlansFrndlyName.csv | ForEach-Object {
 $ServicePlanHash[$_.ServicePlanId] = $_.ServicePlanFriendlyNames
}


#Process all groups
$RequiredProperties=@('DisplayName','Description','Id','AssignedLicenses','LicenseProcessingState')
Get-MgBetaGroup -All -Filter "assignedLicenses/`$count ne 0" -Property $RequiredProperties -ConsistencyLevel eventual -CountVariable Records | select $RequiredProperties | ForEach-Object {
 
 $Count++
 $AssignedLicenseNames=@()
 $AssignedLicenseSKUs=@()
 $AssignedLicense_FrndlyNames=@()
 $DisabledServices=""
 $DisplayName=$_.DisplayName
 $Description=$_.Description
 $Id=$_.Id
 $AssignedLicenses=($_.AssignedLicenses)
 $State=($_.LicenseProcessingState).State
 Write-Progress -Activity "`n     Licensed group processed: $Count - $DisplayName"
 #Processing and converting license to it's friendly name
 foreach($License in $AssignedLicenses)
 {
  $SKU=$License.SkuId
  $SkuName=$SkuHashtable[$SKU]
  $FriendlyName=Convert-FrndlyName -InputIds $SkuName
  $AssignedLicenseNames += $SkuHashtable[$SKU]
  $AssignedLicense_FrndlyNames += $FriendlyName
  $AssignedLicenseSKUs +=$SKU
  $DisabledPlans=$License.DisabledPlans

  #Checking for disabled plans
  $ServicePlanNames=@()
   $DisabledServicePlans=@()
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
   if($DisabledPlans -eq "")
   {
    $DisabledPlans="None"
   }
   $DisabledServicePlans = $SkuName +"[" +$DisabledPlans +"]"
   If($DisabledServices -ne "")
   {
    $DisabledServices= $DisabledServices +","
   }
   $DisabledServices += $DisabledServicePlans
   
 }
 $AssignedLicenseNames=$AssignedLicenseNames -join ","
 $AssignedLicenseSKUs=$AssignedLicenseSKUs -join ","
 $AssignedLicense_FrndlyNames=$AssignedLicense_FrndlyNames -join ","

 $OwnersCount=Get-MgBetaGroupOwnerCount -GroupId $Id -ConsistencyLevel Eventual
 $MembersCount=Get-MgBetaGroupMemberCount -GroupId $Id -ConsistencyLevel eventual
 $GroupUserCountTotal=$OwnersCount+$MembersCount
 $UsersWithLicenseAssignmentError=(Get-MgBetaGroupMemberWithLicenseError -GroupId $Id).Count

 $ExportResult=[PSCustomObject]@{'Display Name'=$DisplayName;'Description'=$Description;'Assigned Licenses'=$AssignedLicenseNames;'Assigned Licenses(Friendly-name)'=$AssignedLicense_FrndlyNames;'Disabled Services'=$DisabledServices;'Group User Count Total'=$GroupUserCountTotal;'Member Count'=$MembersCount;'Owner Count'=$OwnersCount;'Users With License Assignment Error'=$UsersWithLicenseAssignmentError;'State'=$State;'Assigned License SKUs'=$AssignedLicenseSKUs;'Group Id'=$Id}
 $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
}

#Open output file after execution 
 If($OutputCount -eq 0)
 {
  Write-Host No data found for the given criteria
 }
 else
 {
  Write-Host `nThe output file contains $Count groups.
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
 