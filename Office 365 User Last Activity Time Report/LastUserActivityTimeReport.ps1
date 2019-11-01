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
    [switch]$MFA
    
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
 #Check for MSOnline module
 $Modules=Get-Module -Name MSOnline -ListAvailable 
 if($Modules.count -eq 0)
 {
  Write-Host  Please install MSOnline module using below command: `nInstall-Module MSOnline  -ForegroundColor yellow 
  Exit
 }
 #Connect AzureAD and Exchange Online from PowerShell 
 Get-PSSession | Remove-PSSession 

 #Get friendly name of license plan from external file 
 $FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData


 #Set output file 
 $ExportCSV=".\LastAccessTimeReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"

 #Authentication using MFA
 if($MFA.IsPresent)
 {
  $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
  If ($MFAExchangeModule -eq $null)
  {
   Write-Host  `nPlease install Exchange Online MFA Module.  -ForegroundColor yellow
   
   Write-Host You can install module using below blog : `nLink `nOR you can install module directly by entering "Y"`n
   $Confirm= Read-Host Are you sure you want to install module directly? [Y] Yes [N] No
   if($Confirm -match "[yY]")
   {
     Write-Host Yes
     Start-Process "iexplore.exe" "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"
   }
   else
   {
    Start-Process 'https://o365reports.com/2019/03/23/export-dynamic-distribution-group-members-to-csv/'
    Exit
   }
   $Confirmation= Read-Host Have you installed Exchange Online MFA Module? [Y] Yes [N] No
   if($Confirmation -match "[yY]")
   {
    $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
    If ($MFAExchangeModule -eq $null)
    {
     Write-Host Exchange Online MFA module is not available -ForegroundColor red
     Exit
    }
   }
   else
   { 
    Write-Host Exchange Online PowerShell Module is required
    Start-Process 'https://o365reports.com/2019/03/23/export-dynamic-distribution-group-members-to-csv/'
    Exit
   }   
  }
  
  #Importing Exchange MFA Module
  . "$MFAExchangeModule"
  Write-Host Enter credential in prompt to connect to Exchange Online
  Connect-EXOPSSession -WarningAction SilentlyContinue
  Write-Host Connected to Exchange Online
  Write-Host `nEnter credential in prompt to connect to MSOnline
  #Importing MSOnline Module
  Connect-MsolService | Out-Null
  Write-Host Connected to MSOnline `n`nReport generation in progress...
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
  Connect-MsolService -Credential $credential 
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
  Import-PSSession $Session -CommandName Get-Mailbox,Get-MailboxStatistics -FormatTypeName * -AllowClobber | Out-Null
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
  Get-Mailbox -ResultSize Unlimited | Where{$_.DisplayName -notlike "Discovery Search Mailbox"} | ForEach-Object{ 
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
  Write-Host "Detailed report available in: $ExportCSV" 
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