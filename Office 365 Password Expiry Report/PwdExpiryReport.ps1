<#
=============================================================================================
Name:           Export Office 365 Users’ Last Password Change Date to CSV 
website:        o365reports.com
Script by:      O365Reports Team
For detailed Script execution: https://o365reports.com/2020/02/17/export-office-365-users-last-password-change-date-to-csv
============================================================================================
#>
Param 
( 
    [Parameter(Mandatory = $false)] 
    [switch]$PwdNeverExpires, 
    [switch]$PwdExpired, 
    [switch]$LicensedUserOnly, 
    [int]$SoonToExpire, 
    [int]$RecentPwdChanges,
    [switch]$EnabledUsersOnly,
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
$PwdPolicy=@{}
$Results=@()  
$UserCount=0 
$PrintedUser=0 

#Output file declaration 
$ExportCSV=".\PasswordExpiryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 

#Getting Password policy for the domain
$Domains=Get-MsolDomain   #-Status Verified
foreach($Domain in $Domains)
{ 
 #Check for federated domain
 if($Domain.Authentication -eq "Federated")
 {
  $PwdValidity=0
 }
 else
 {
  $PwdValidity=(Get-MsolPasswordPolicy -DomainName $Domain.Name -ErrorAction SilentlyContinue ).ValidityPeriod
  if($PwdValidity -eq $null)
  {                                 
   $PwdValidity=90
  }
 }
 $PwdPolicy.Add($Domain.name,$PwdValidity)
}
 Write-Host Generating report...
#Loop through each user 
Get-MsolUser -All | foreach{ 
 $UPN=$_.UserPrincipalName
 $DisplayName=$_.DisplayName
 [boolean]$Federated=$false
 $UserCount++
 #Remove external users
 if($UPN -like "*#EXT#*")
 {
  return
 }

 $PwdLastChange=$_.LastPasswordChangeTimestamp
 $PwdNeverExpire=$_.PasswordNeverExpires
 $LicenseStatus=$_.isLicensed
 $Print=0
 Write-Progress -Activity "`n     Processed user count: $UserCount "`n"  Currently Processing: $DisplayName"
 if($LicenseStatus -eq $true)
 {
  $LicenseStatus="Licensed"
 }
 else
 {
  $LicenseStatus="Unlicensed"
 }
 
 if($_.BlockCredential -eq $true)
 {
  $AccountStatus="Disabled"
 }
 else
 {
  $AccountStatus="Enabled"
 }

 #Finding password validity period for user
 $UserDomain= $UPN -Split "@" | Select-Object -Last 1 
 $PwdValidityPeriod=$PwdPolicy[$UserDomain]

 #Check for Pwd never expires set from pwd policy
 if([int]$PwdValidityPeriod -eq 2147483647)
 {
  $PwdNeverExpire=$true
  $PwdExpireIn="Never Expires"
  $PwdExpiryDate="-"
  $PwdExpiresIn="-"
 }
 elseif($PwdValidityPeriod -eq 0) #Users from federated domain
 {
  $Federated=$true
  $PwdExpireIn="Insufficient data in O365"
  $PwdExpiryDate="-"
  $PwdExpiresIn="-"
 }
 elseif($PwdNeverExpire -eq $False) #Check for Pwd never expires set from Set-MsolUser
 {
  $PwdExpiryDate=$PwdLastChange.AddDays($PwdValidityPeriod)
  $PwdExpiresIn=(New-TimeSpan -Start (Get-Date) -End $PwdExpiryDate).Days
  if($PwdExpiresIn -gt 0)
  {
   $PwdExpireIn= "in $PwdExpiresIn days"
  }
  elseif($PwdExpiresIn -lt 0)
  {
   #Write-host `n $PwdExpiresIn
   $PwdExpireIn =$PwdExpiresIn * (-1)
   #Write-Host ************$pwdexpiresin
   $PwdExpireIn="$PwdExpireIn days ago"
  }
  else
  {
   $PwdExpireIn="Today"
  }
 }
 else
 {
  $PwdExpireIn="Never Expires"
  $PwdExpiryDate="-"
  $PwdExpiresIn="-"
 }

 #Calculating Password since last set
 $PwdSinceLastSet=(New-TimeSpan -Start $PwdLastChange).Days

 #Filter for enabled users
 if(($EnabledUsersOnly.IsPresent) -and ($AccountStatus -eq "Disabled"))
 {
  return
 }

 #Filter for user with Password nerver expires
 if(($PwdNeverExpires.IsPresent) -and ($PwdNeverExpire -eq $false))
 {
  return
 }
 
 #Filter for password expired users
 if(($pwdexpired.IsPresent) -and (($PwdExpiresIn -ge 0) -or ($PwdExpiresIn -eq "-")))
 { 
  return
 }

 #Filter for licensed users
 if(($LicensedUserOnly.IsPresent) -and ($LicenseStatus -eq "Unlicensed"))
 {
  return
 }

 #Filter for soon to expire pwd users
 if(($SoonToExpire -ne "") -and (($PwdExpiryDate -eq "-") -or ([int]$SoonToExpire -lt $PwdExpiresIn) -or ($PwdExpiresIn -lt 0)))
 { 
  return
 }

 #Filter for recently password changed users
 if(($RecentPwdChanges -ne "") -and ($PwdSinceLastSet -gt $RecentPwdChanges))
 {
  return
 }

 if($Federated -eq $true)
 {
  $PwdExpiryDate="Insufficient data in O365"
  $PwdExpiresIn="Insufficient data in O365"
 }

 $PrintedUser++ 
 
 #Export result to csv
 $Result=@{'Display Name'=$DisplayName;'User Principal Name'=$upn;'Pwd Last Change Date'=$PwdLastChange;'Days since Pwd Last Set'=$PwdSinceLastSet;'Pwd Expiry Date'=$PwdExpiryDate;'Days since Expiry(-) / Days to Expiry(+)'=$PwdExpiresIn ;'Friendly Expiry Time'=$PwdExpireIn;'License Status'=$LicenseStatus;'Account Status'=$AccountStatus}
 $Results= New-Object PSObject -Property $Result  
 $Results | Select-Object 'Display Name','User Principal Name','Pwd Last Change Date','Days since Pwd Last Set','Pwd Expiry Date','Friendly Expiry Time','License Status','Days since Expiry(-) / Days to Expiry(+)','Account Status' | Export-Csv -Path $ExportCSV -Notype -Append 
}

If($UserCount -eq 0)
{
 Write-Host No records found
}
else
{
 Write-Host `nThe output file contains $PrintedUser users.
 if((Test-Path -Path $ExportCSV) -eq "True") 
 {
  Write-Host `nThe Output file available in $ExportCSV -ForegroundColor Green
   $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$ExportCSV"   
  } 
 }
}
