<#
=============================================================================================
Name:           Export Office 365 users real last activity time report
Version:        3.0
Website:        o365reports.com
Script by:      O365Reports Team

Script Highlights: 
~~~~~~~~~~~~~~~~~

1.Reports the user’s activity time based on the user’s last action time(LastUserActionTime). 
2.Exports result to CSV file. 
3.Result can be filtered based on inactive days. 
4.You can filter the result based on user/mailbox type. 
5.Result can be filtered to list never logged in mailboxes alone. 
6.You can filter the result based on licensed user.
7.Shows result with the user’s administrative roles in the Office 365. 
8.The assigned licenses column will show you the user-friendly-name like ‘Office 365 Enterprise E3’ rather than ‘ENTERPRISEPACK’. 
9.The script can be executed with MFA enabled account. 
10.The script is scheduler friendly. i.e., credentials can be passed as a parameter instead of saving inside the script. 

For detailed script execution:  https://o365reports.com/2019/06/18/export-office-365-users-real-last-logon-time-report-csv/#
============================================================================================
#>

#Accept input parameter
Param
(
    [Parameter(Mandatory = $false)]
    [string]$MBNamesFile,
    [int]$InactiveDays,
    [switch]$UserMailboxOnly,
    [switch]$LicensedUserOnly,
    [switch]$ReturnNeverLoggedInMBOnly,
    [string]$UserName,
    [string]$Password,
    [switch]$FriendlyTime,
    [switch]$NoMFA
    
)

Function Get_LastLogonTime
{
 $MailboxStatistics=Get-MailboxStatistics -Identity $upn
 $LastActionTime=$MailboxStatistics.LastUserActionTime
 $LastActionTimeUpdatedOn=$MailboxStatistics.LastUserActionUpdateTime
 $RolesAssigned="" 
 Write-Progress -Activity "`n     Processed mailbox count: $MBUserCount "`n"  Currently Processing: $DisplayName" 
 
 #Retrieve lastlogon time and then calculate Inactive days 
 if($LastActionTime -eq $null)
 { 
   $LastActionTime ="Never Logged In" 
   $InactiveDaysOfUser="-" 
 } 
 else
 { 
   $InactiveDaysOfUser= (New-TimeSpan -Start $LastActionTime).Days
   #Convert Last Action Time to Friendly Time
   if($friendlyTime.IsPresent) 
   {
    $FriendlyLastActionTime=ConvertTo-HumanDate ($LastActionTime)
    $friendlyLastActionTime="("+$FriendlyLastActionTime+")"
    $LastActionTime="$LastActionTime $FriendlyLastActionTime" 
   }
 }
  #Convert Last Action Time Update On to Friendly Time
 if(($friendlyTime.IsPresent) -and ($LastActionTimeUpdatedOn -ne $null))
 {
  $FriendlyLastActionTimeUpdatedOn= ConvertTo-HumanDate ($LastActionTimeUpdatedOn)
  $FriendlyLastActionTimeUpdatedOn="("+$FriendlyLastActionTimeUpdatedOn+")"
  $LastActionTimeUpdatedOn="$LastActionTimeUpdatedOn $FriendlyLastActionTimeUpdatedOn" 
 }
 elseif($LastActionTimeUpdatedOn -eq $null)
 {
  $LastActionTimeUpdatedOn="-"
 }
 
 #Get licenses assigned to mailboxes 
 $User=(Get-MsolUser -UserPrincipalName $upn) 
 $Licenses=$User.Licenses.AccountSkuId 
 $AssignedLicense="" 
 $Count=0
 
 
 if($Licenses.count -eq 0) 
 { 
  $AssignedLicense="No License Assigned" 
 }  
 #Convert license plan to friendly name 
 else
 {
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
  }

 #Inactive days based filter 
 if($InactiveDaysOfUser -ne "-"){ 
 if(($InactiveDays -ne "") -and ([int]$InactiveDays -gt $InactiveDaysOfUser)) 
 { 
  return
 }} 

 #Filter result based on user mailbox 
 if(($UserMailboxOnly.IsPresent) -and ($MBType -ne "UserMailbox"))
 { 
  return
 } 

 #Never Logged In user
 if(($ReturnNeverLoggedInMBOnly.IsPresent) -and ($LastActionTime -ne "Never Logged In"))
 {
  return
 }

 #Filter result based on license status
 if(($LicensedUserOnly.IsPresent) -and ($AssignedLicense -eq "No License Assigned"))
 {
  return
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
 $Result=@{'UserPrincipalName'=$upn;'DisplayName'=$DisplayName;'LastUserActionTime'=$LastActionTime;'LastActionTimeUpdatedOn'=$LastActionTimeUpdatedOn;'CreationTime'=$CreationTime;'InactiveDays'=$InactiveDaysOfUser;'MailboxType'=$MBType; 'AssignedLicenses'=$AssignedLicense;'Roles'=$RolesAssigned} 
 $Output= New-Object PSObject -Property $Result 
 $Output | Select-Object UserPrincipalName,DisplayName,LastUserActionTime,LastActionTimeUpdatedOn,InactiveDays,CreationTime,MailboxType,AssignedLicenses,Roles | Export-Csv -Path $ExportCSV -Notype -Append
} 


Function main()
{
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
   Import-Module ExchangeOnlineManagement
  } 
  else 
  { 
   Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 
 #Check for Azure AD module
 $Module = Get-Module MsOnline -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host MSOnline module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install the module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing MSOnline PowerShell module"
   Install-Module MSOnline -Repository PSGallery -AllowClobber -Force
   Import-Module MSOnline
  } 
  else 
  { 
   Write-Host MSOnline module is required to generate the report.Please install module using Install-Module MSOnline cmdlet. 
   Exit
  }
 }

 #Authentication using non-MFA
 if($NoMFA.IsPresent)
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
  Write-Host "Connecting Azure AD..."
  Connect-MsolService -Credential $Credential | Out-Null
  Write-Host "Connecting Exchange Online PowerShell..."
  Connect-ExchangeOnline -Credential $Credential
 }
 #Connect to Exchange Online and AzureAD module using MFA 
 else
 {
  Write-Host "Connecting Exchange Online PowerShell..."
  Connect-ExchangeOnline
  Write-Host "Connecting Azure AD..."
  Connect-MsolService | Out-Null
 }

 #Friendly DateTime conversion
 if($friendlyTime.IsPresent)
 {
  If(((Get-Module -Name PowerShellHumanizer -ListAvailable).Count) -eq 0)
  {
   Write-Host Installing PowerShellHumanizer for Friendly DateTime conversion
   Install-Module -Name PowerShellHumanizer
  }
 }

 $Result=""  
 $Output=@() 
 $MBUserCount=0 

 #Get friendly name of license plan from external file 
 $FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData


 #Set output file 
 $ExportCSV=".\LastAccessTimeReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"

 #Check for input file
 if([string]$MBNamesFile -ne "") 
 { 
  #We have an input file, read it into memory 
  $Mailboxes=@()
  $Mailboxes=Import-Csv -Header "MBIdentity" $MBNamesFile
  foreach($item in $Mailboxes)
  {
   $MBDetails=Get-Mailbox -Identity $item.MBIdentity
   $upn=$MBDetails.UserPrincipalName 
   $CreationTime=$MBDetails.WhenCreated
   $DisplayName=$MBDetails.DisplayName 
   $MBType=$MBDetails.RecipientTypeDetails
   $MBUserCount++
   Get_LastLogonTime    
  }
 }

 #Get all mailboxes from Office 365
 else
 {
  Write-Progress -Activity "Getting Mailbox details from Office 365..." -Status "Please wait." 
  Get-Mailbox -ResultSize Unlimited | Where{$_.DisplayName -notlike "Discovery Search Mailbox"} | ForEach { 
  $upn=$_.UserPrincipalName 
  $CreationTime=$_.WhenCreated
  $DisplayName=$_.DisplayName 
  $MBType=$_.RecipientTypeDetails
  $MBUserCount++
  Get_LastLogonTime
  } 
 }

 #Open output file after execution 
 Write-Host `nScript executed successfully
 if((Test-Path -Path $ExportCSV) -eq "True")
 {
  Write-Host `n" Detailed report available in:" -NoNewline -ForegroundColor Yellow
  Write-Host $ExportCSV
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n 
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
}
 . main