<#
=============================================================================================
Name:           Get License Expiry Date report
Version:        3.0
Website:        o365reports.com

ChangeLog
~~~~~~~~~

   V1 - Initial version (04/03/2020)
   V2 - Minor changes (10/06/2023)
   V3 - Migrated from MSOnline module to MS Graph (9/13/2024)


Script Highlights: 
~~~~~~~~~~~~~~~~~~

1.Exports Office 365 license expiry date with ‘next lifecycle activity date’. 
2.Exports report to the CSV file. 
3.Result can be filtered based on subscription type like Purchased, Trial and Free subscription 
4.Result can be filtered based on subscription status like Enabled, Expired, Disabled, etc. 
5.Subscription name is shown as user-friendly-name like ‘Office 365 Enterprise E3’ rather than ‘ENTERPRISEPACK’. 
6.The script can be executed with MFA enabled account too. 
7.The script is scheduler friendly.
8.Supports certificate based authentication (CBA) too.

For detailed script execution:  https://o365reports.com/2020/03/04/export-office-365-license-expiry-date-report-powershell/
============================================================================================
#>
Param 
( 
    [Parameter(Mandatory = $false)] 
    [switch]$Trial, 
    [switch]$Free, 
    [switch]$Purchased, 
    [Switch]$Expired, 
    [Switch]$Active,
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

Connect_MgGraph

$Result=""   
$Results=@()  
$Print=0
$ShowAllSubscription=$False
$PrintedOutput=0

#Output file declaration 
$Location=Get-Location
$ExportCSV="$Location\LicenseExpiryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 

#Check for filters
if((!($Trial.IsPresent)) -and (!($Free.IsPresent)) -and (!($Purchased.IsPresent)) -and (!($Expired.IsPresent)) -and (!($Active.IsPresent)))
{
 $ShowAllSubscription=$true
}

#FriendlyName list for license plan 
$FriendlyNameHash=@()
$FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData 


#Get next lifecycle date
$ExpiryDateHash=@{}
$LifeCycleDateInfo=(Invoke-MgGraphRequest -Uri https://graph.microsoft.com/V1.0/directory/subscriptions -Method Get).Value
foreach($Date in $LifeCycleDate)
{
 $ExpiryDateHash.Add($Date.skuId,$Date.nextLifeCycleDateTime)
}

#Get available subscriptions in the tenant
$Subscriptions= Get-MgBetaSubscribedSku -All | foreach{
 $SubscriptionName=$_.SKUPartNumber
 $SkuId=$_.SkuId
 $ConsumedUnits=$_.ConsumedUnits
 $MoreSkuDetails=$LifeCycleDateInfo | Where {$_.skuId -eq $SkuId}
 $SubscribedOn=$MoreSkuDetails.createdDateTime
 $Status=$MoreSkuDetails.status
 $TotalLicenses=$MoreSkuDetails.totalLicenses
 $ExpiryDate=$MoreSkuDetails.nextLifeCycleDateTime
 $RemainingUnits=$TotalLicenses - $ConsumedUnits
 $Print=0

 #Convert Skuid to friendly name  
 $EasyName=$FriendlyNameHash[$SubscriptionName] 
 $EasyName
 if(!($EasyName)) 
 {
  $NamePrint=$SubscriptionName
 } 
 else 
 {
  $NamePrint=$EasyName
 } 
 
 #Convert Subscribed date to friendly subscribed date
 $SubscribedDate=(New-TimeSpan -Start $SubscribedOn -End (Get-Date)).Days
 if($SubscribedDate -eq 0)
 {
  $SubscribedDate="Today"
 }
 else
 {
  $SubscribedDate="$SubscribedDate days ago"
 }
 $SubscribedDate="(" + $SubscribedDate + ")"
 $SubscribedDate="$SubscribedOn $SubscribedDate"

 #Determine subscription type
  if(($SubscriptionName -like "*Free*") -and ($ExpiryDate -eq $null))
  {
   $SubscriptionType="Free"
  }
  elseif($ExpiryDate -eq $null)
  {
   $SubscriptionType="Trial"
  }
 else
 {
  $SubscriptionType="Purchased"
 }
 
 #Friendly Expiry Date
 if($ExpiryDate -ne $null)
 {
  $FriendlyExpiryDate=(New-TimeSpan -Start (Get-Date) -End $ExpiryDate).Days
  if($Status -eq "Enabled")
  {
   $FriendlyExpiryDate="Will expire in $FriendlyExpiryDate days"
  }
  elseif($Status -eq "Warning")
  {
   $FriendlyExpiryDate="Expired.Will suspend in $FriendlyExpiryDate days"
  }
  elseif($Status -eq "Suspended")
  {
   $FriendlyExpiryDate="Expired.Will delete in $FriendlyExpiryDate days"
  }
  elseif($Status -eq "LockedOut")
  {
   $FriendlyExpiryDate="Subscription is locked.Please contact Microsoft"
  }
 }
 else
 {
  $ExpiryDate="-"
  $FriendlyExpiryDate="Never Expires"
 }
 
 #Check for filters
 if($ShowAllSubscription -eq $true)
 {
  $Print=1
 }
 else
 {
  if(($Trial.IsPresent) -and ($SubscriptionType -eq "Trial"))
  {
   $Print=1
  }
  if(($Free.IsPresent) -and ($SubscriptionType -eq "Free"))
  {
   $Print=1
  }
  if(($Purchased.IsPresent) -and ($SubscriptionType -eq "Purchased"))
  {
   $Print=1
  }
  if(($Expired.IsPresent) -and ($Status -ne "Enabled"))
  {
   $Print=1
  }
  if(($Active.IsPresent) -and ($Status -eq "Enabled"))
  {
   $Print=1
  }
 }
 


 #Export result to csv
 if($Print -eq 1)
 {
  $PrintedOutput++
  $Result=@{'Subscription Name'=$SubscriptionName;'SKU Id'=$SkuId;'Friendly Subscription Name'=$NamePrint;'Subscribed Date'=$SubscribedDate;'Total Units'=$TotalLicenses;'Consumed Units'=$ConsumedUnits;'Remaining Units'=$RemainingUnits;'License Expiry Date/Next LifeCycle Activity Date'=$ExpiryDate;'Friendly Expiry Date'=$FriendlyExpiryDate;'Subscription Type'=$SubscriptionType;'Status'=$Status}
  $Results= New-Object PSObject -Property $Result  
  $Results | Select-Object 'Subscription Name','Friendly Subscription Name','Subscribed Date','Total Units','Consumed Units','Remaining Units','Subscription Type','License Expiry Date/Next LifeCycle Activity Date','Friendly Expiry Date','Status','SKU Id' | Export-Csv -Path $ExportCSV -Notype -Append 
 }
}

#Open output file after execution 
if((Test-Path -Path $ExportCSV) -eq "True") 
{
 Write-Host ""
 Write-Host " Office 365 license expiry report available in:"  -NoNewline -ForegroundColor Yellow
 Write-Host $ExportCSV 
 Write-Host ""
 Write-Host " The Output file contains:" $PrintedOutput subscriptions  
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n 
 $Prompt = New-Object -ComObject wscript.shell 
 $UserInput = $Prompt.popup("Do you want to open output files?",` 
 0,"Open Files",4) 
 If ($UserInput -eq 6) 
 { 
  Invoke-Item "$ExportCSV" 
 } 
}
else
{
 Write-Host No subscription found.
}