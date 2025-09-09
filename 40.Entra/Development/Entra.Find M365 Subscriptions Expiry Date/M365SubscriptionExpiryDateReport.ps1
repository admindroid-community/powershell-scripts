<#
=============================================================================================
Name:           Get Microsoft 365 Subscription Expiry Date Report
Version:        1.0


Script Highlights: 
~~~~~~~~~~~~~~~~~~

1.Exports Office 365 license expiry date with ‘next lifecycle activity date’. 
2.Exports report to the CSV file. 
3.Result can be filtered based on soon to be expire license. ie, Licenses that are about to expire in 30 days.
4.Allows to filter out 'Purchased subscriptions' alone.
5.Subscription name is shown as user-friendly-name like ‘Office 365 Enterprise E3’ rather than ‘ENTERPRISEPACK’. 
6.The script can be executed with MFA enabled account too. 
7.The script supports Certificate based authentication too.
8.The script is scheduler friendly. 


============================================================================================
#>
Param 
( 
    [Parameter(Mandatory = $false)] 

    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [switch]$PurchasedSubscriptionsOnly, 
    [int]$SoonToExpireInDays
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

$PrintedOutput=0

#Output file declaration 
$Location=Get-Location
$ExportCSV="$Location\M365SubscriptionExpiryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 

#Check for filters


#FriendlyName list for license plan 
$FriendlyNameHash=@()
$FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData 


#Get next lifecycle date

$SubscriptionDetails=(Invoke-MgGraphRequest -Uri https://graph.microsoft.com/V1.0/directory/subscriptions -Method Get).Value


#Get available subscriptions in the tenant
foreach ($Subscription in $SubscriptionDetails)
{
 $SubscriptionName=$Subscription.SKUPartNumber
 $SkuId=$Subscription.SkuId
 $SubscribedOn=$Subscription.createdDateTime
 $Status=$Subscription.status
 $TotalLicenses=$Subscription.totalLicenses
 $ExpiryDate=$Subscription.nextLifeCycleDateTime
 $Print=1

 #Determine subscription type
  if(($ExpiryDate -eq $null))
  {
   $SubscriptionType="Free/Trial"
  }
  else
  {
   $SubscriptionType="Purchased"
  }
  
  #Filter for purchased licenses
  if(($PurchasedSubscriptionsOnly.IsPresent) -and ($SubscriptionType -ne "Purchased"))
  {
   $Print=0
  }


 #Convert Skuid to friendly name  
 $EasyName=$FriendlyNameHash[$SubscriptionName] 
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
 
 #Filter for soon-to-expire subscriptions
 if(($SoonToExpireInDays -ne "") -and ($ExpiryDate -ne $null) -and (($SoonToExpireInDays -lt $FriendlyExpiryDate)))
 {
  $Print=0
 }


 #Export result to csv

 if($Print -eq 1)
 {
  $PrintedOutput++
  $Result=@{'Subscription Name'=$SubscriptionName;'SKU Id'=$SkuId;'Friendly Subscription Name'=$NamePrint;'Subscribed Date'=$SubscribedDate;'Total Units'=$TotalLicenses;'License Expiry Date/Next LifeCycle Activity Date'=$ExpiryDate;'Friendly Expiry Date'=$FriendlyExpiryDate;'Subscription Type'=$SubscriptionType;'Status'=$Status}
  $Results= New-Object PSObject -Property $Result  
  $Results | Select-Object 'Subscription Name','Friendly Subscription Name','Subscribed Date','Total Units','Subscription Type','License Expiry Date/Next LifeCycle Activity Date','Friendly Expiry Date','Status','SKU Id' | Export-Csv -Path $ExportCSV -Notype -Append 
 }
}

#Open output file after execution 
if((Test-Path -Path $ExportCSV) -eq "True") 
{
 Write-Host ""
 Write-Host " Office 365 license expiry report available in:"  -NoNewline -ForegroundColor Yellow
 Write-Host $ExportCSV 
 Write-Host ""
 Write-Host " The Output file contains:" -NoNewline -ForegroundColor Yellow
 Write-Host $PrintedOutput subscriptions  
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