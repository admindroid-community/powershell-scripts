<#
=============================================================================================
Name:           Get License Expiry Date report
Version:        2.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~~

1.Exports Office 365 license expiry date with ‘next lifecycle activity date’. 
2.Exports report to the CSV file. 
3.Result can be filtered based on subscription type like Purchased, Trial and Free subscription 
4.Result can be filtered based on subscription status like Enabled, Expired, Disabled, etc. 
5.Subscription name is shown as user-friendly-name like ‘Office 365 Enterprise E3’ rather than ‘ENTERPRISEPACK’. 
6.The script can be executed with MFA enabled account too. 
7.The script is scheduler friendly. i.e., credentials can be passed as a parameter instead of saving inside the script. 

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
    [string]$UserName,  
    [string]$Password 
) 

#Check for MSOnline module 
$Module=Get-Module -Name MSOnline -ListAvailable  
if($Module.count -eq 0) 
{ 
 Write-Host MSOnline module is not available  -ForegroundColor yellow  
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
 if($Confirm -match "[yY]") 
 { 
  Install-Module MSOnline 
  Import-Module MSOnline
 } 
 else 
 { 
  Write-Host MSOnline module is required to connect AzureAD.Please install module using Install-Module MSOnline cmdlet. 
  Exit
 }
} 
 
#Storing credential in script for scheduling purpose/ Passing credential as parameter  
if(($UserName -ne "") -and ($Password -ne ""))  
{  
 $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force  
 $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword  
 Connect-MsolService -Credential $credential 
}  
else  
{  
 Connect-MsolService | Out-Null  
} 

$Result=""   
$Results=@()  
$Print=0
$ShowAllSubscription=$False
$PrintedOutput=0

#Output file declaration 
$ExportCSV=".\LicenseExpiryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 

#Check for filters
if((!($Trial.IsPresent)) -and (!($Free.IsPresent)) -and (!($Purchased.IsPresent)) -and (!($Expired.IsPresent)) -and (!($Active.IsPresent)))
{
 $ShowAllSubscription=$true
}

#FriendlyName list for license plan 
$FriendlyNameHash=@()
$FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData 

#Get available subscriptions in the tenant
$Subscriptions= Get-MsolSubscription | foreach{
 $SubscriptionName=$_.SKUPartNumber
 $SubscribedOn=$_.DateCreated
 $ExpiryDate=$_.NextLifeCycleDate
 $Status=$_.Status
 $TotalLicenses=$_.TotalLicenses
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
 If($_.IsTrial -eq $False)
 {
  if(($SubscriptionName -like "*Free*") -or ($ExpiryDate -eq $null))
  {
   $SubscriptionType="Free"
  }
  else
  {
   $SubscriptionType="Purchased"
  }
 }
 else
 {
  $SubscriptionType="Trial"
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
  $Result=@{'Subscription Name'=$SubscriptionName;'Friendly Subscription Name'=$NamePrint;'Subscribed Date'=$SubscribedDate;'Total Licenses'=$TotalLicenses;'License Expiry Date/Next LifeCycle Activity Date'=$ExpiryDate;'Friendly Expiry Date'=$FriendlyExpiryDate;'Subscription Type'=$SubscriptionType;'Status'=$Status}
  $Results= New-Object PSObject -Property $Result  
  $Results | Select-Object 'Subscription Name','Friendly Subscription Name','Subscribed Date','Total Licenses','Subscription Type','License Expiry Date/Next LifeCycle Activity Date','Friendly Expiry Date','Status' | Export-Csv -Path $ExportCSV -Notype -Append 
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