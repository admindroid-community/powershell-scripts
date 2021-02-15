<#
=============================================================================================
Name:           Export Office 365 user last logon time report
Description:    This script exports Office 365 users' last logon time CSV
Version:        3.0
website:        o365reports.com
Script by:      O365Reports Team
For detailed Script execution: https://o365reports.com/2019/03/07/export-office-365-users-last-logon-time-csv/
============================================================================================
#>
#Accept input parameter
Param
(
    [Parameter(Mandatory = $false)]
    [int]$InactiveDays,
    [switch]$UserMailboxOnly,
    [switch]$ReturnNeverLoggedInMB,
    [string]$UserName,
    [string]$Password,
    [switch]$MFA

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

#Clear session
Get-PSSession | Remove-PSSession

#Get friendly name of license plan from external file
$FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData


#Set output file
$ExportCSV=".\LastLogonTimeReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"

 #Check for EXO v2 module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell V2 module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
  } 
  else 
  { 
   Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 
 
 #Connect Exchange Online with MFA
 if($MFA.IsPresent)
 {
  Write-Host Connecting Exchange Online...
  Connect-ExchangeOnline | Out-Null
  Write-Host Connecting to Office 365...
  Connect-MsolService | Out-Null
 }

 #Authentication using non-MFA
 else
 {
  #Storing credential in script for scheduling purpose/ Passing credential as parameter
  if(($UserName -ne "") -and ($Password -ne ""))
  {
   $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
   $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  }
  else
  {
   $Credential=Get-Credential -Credential $null
  }
  Write-Host Connecting Exchange Online...
  Connect-ExchangeOnline -Credential $Credential
  Write-Host Connecting to Office 365...
  Connect-MsolService -Credential $Credential
 }

$Result=""
$Output=@()
$MBUserCount=0
$OutputCount=0


Get-Mailbox -ResultSize Unlimited | Where{$_.DisplayName -notlike "Discovery Search Mailbox"} | ForEach-Object{
 $upn=$_.UserPrincipalName
 $CreationTime=$_.WhenCreated
 $LastLogonTime=(Get-MailboxStatistics -Identity $upn).lastlogontime
 $DisplayName=$_.DisplayName
 $MBType=$_.RecipientTypeDetails
 $Print=1
 $MBUserCount++
 $RolesAssigned=""
 Write-Progress -Activity "`n     Processed mailbox count: $MBUserCount "`n"  Currently Processing: $DisplayName"

 #Retrieve lastlogon time and then calculate Inactive days
 if($LastLogonTime -eq $null)
 {
   $LastLogonTime ="Never Logged In"
   $InactiveDaysOfUser="-"
 }
 else
 {
   $InactiveDaysOfUser= (New-TimeSpan -Start $LastLogonTime).Days
 }

 #Get licenses assigned to mailboxes
 $User=(Get-MsolUser -UserPrincipalName $upn)
 $Licenses=$User.Licenses.AccountSkuId
 $AssignedLicense=""
 $Count=0

 #Convert license plan to friendly name
 foreach($License in $Licenses)
 {
    $Count++
    $LicenseItem= $License -Split ":" | Select-Object -Last 1
    $EasyName=$FriendlyNameHash[$LicenseItem]
    if(!($EasyName))
    {$NamePrint=$LicenseItem}
    else
    {$NamePrint=$EasyName}
    $AssignedLicense=$AssignedLicense+$NamePrint
    if($count -lt $licenses.count)
    {
      $AssignedLicense=$AssignedLicense+","
    }
 }
 if($Licenses.count -eq 0)
 {
  $AssignedLicense="No License Assigned"
 }

 #Inactive days based filter
 if($InactiveDaysOfUser -ne "-"){
 if(($InactiveDays -ne "") -and ([int]$InactiveDays -gt $InactiveDaysOfUser))
 {
  $Print=0
 }}

 #License assigned based filter
 if(($UserMailboxOnly.IsPresent) -and ($MBType -ne "UserMailbox"))
 {
  $Print=0
 }

 #Never Logged In user
 if(($ReturnNeverLoggedInMB.IsPresent) -and ($LastLogonTime -ne "Never Logged In"))
 {
  $Print=0
 }

 #Get roles assigned to user
 $Roles=(Get-MsolUserRole -UserPrincipalName $upn).Name
 if($Roles.count -eq 0)
 {
  $RolesAssigned="No roles"
 }
 else
 {
  foreach($Role in $Roles)
  {
   $RolesAssigned=$RolesAssigned+$Role
   if($Roles.indexof($role) -lt (($Roles.count)-1))
   {
    $RolesAssigned=$RolesAssigned+","
   }
  }
 }

 #Export result to CSV file
 if($Print -eq 1)
 {
  $OutputCount++
  $Result=@{'UserPrincipalName'=$upn;'DisplayName'=$DisplayName;'LastLogonTime'=$LastLogonTime;'CreationTime'=$CreationTime;'InactiveDays'=$InactiveDaysOfUser;'MailboxType'=$MBType; 'AssignedLicenses'=$AssignedLicense;'Roles'=$RolesAssigned}
  $Output= New-Object PSObject -Property $Result
  $Output | Select-Object UserPrincipalName,DisplayName,LastLogonTime,CreationTime,InactiveDays,MailboxType,AssignedLicenses,Roles | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

#Open output file after execution
Write-Host `nScript executed successfully
if((Test-Path -Path $ExportCSV) -eq "True")
{
 
 Write-Host "Detailed report available in: $ExportCSV"
 
 Write-Host Exported report has $OutputCount mailboxes
 $Prompt = New-Object -ComObject wscript.shell
 $UserInput = $Prompt.popup("Do you want to open output file?",`
 0,"Open Output File",4)
 If ($UserInput -eq 6)
 {
  Invoke-Item "$ExportCSV"
 }
}
Else
{
  Write-Host No mailbox found
}
#Clean up session
Get-PSSession | Remove-PSSession
