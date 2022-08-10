<#
=============================================================================================
Name:           Export Mailbox Permission Report
Website:        o365reports.com
Version:        3.0
For detailed Script execution: https://o365reports.com/2019/03/07/export-mailbox-permission-csv/
============================================================================================
#>
#Accept input paramenters
param(
[switch]$FullAccess,
[switch]$SendAs,
[switch]$SendOnBehalf,
[switch]$UserMailboxOnly,
[switch]$AdminsOnly,
[string]$MBNamesFile,
[string]$UserName,
[string]$Password,
[switch]$NoMFA
)


function Print_Output
{
 #Get admin roles assigned to user 
 if($RolesAssigned -eq "")
 {
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
 }
 #Mailbox type based filter
 if(($UserMailboxOnly.IsPresent) -and ($MBType -ne "UserMailbox"))
 { 
  $Print=0 
 }

 #Admin Role based filter
 if(($AdminsOnly.IsPresent) -and ($RolesAssigned -eq "No roles"))
 { 
  $Print=0 
 }

 #Print Output
 if($Print -eq 1)
 {
  $Result = @{'DisplayName'=$_.Displayname;'UserPrinciPalName'=$upn;'MailboxType'=$MBType;'AccessType'=$AccessType;'UserWithAccess'=$userwithAccess;'Roles'=$RolesAssigned} 
  $Results = New-Object PSObject -Property $Result 
  $Results |select-object DisplayName,UserPrinciPalName,MailboxType,AccessType,UserWithAccess,Roles | Export-Csv -Path $ExportCSV -Notype -Append 
 }
}

#Getting Mailbox permission
function Get_MBPermission
 {
  $upn=$_.UserPrincipalName
  $DisplayName=$_.Displayname
  $MBType=$_.RecipientTypeDetails
  $Print=0
  Write-Progress -Activity "`n     Processed mailbox count: $MBUserCount "`n"  Currently Processing: $DisplayName"

  #Getting delegated Fullaccess permission for mailbox
  if(($FilterPresent -eq 'False') -or ($FullAccess.IsPresent))
  {
   $FullAccessPermissions=(Get-MailboxPermission -Identity $upn | where { ($_.AccessRights -contains "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -match "NT AUTHORITY" -or $_.User -match "S-1-5-21") }).User
   if([string]$FullAccessPermissions -ne "")
   {
    $Print=1
    $UserWithAccess=""
    $AccessType="FullAccess"
    foreach($FullAccessPermission in $FullAccessPermissions)
    {
     $UserWithAccess=$UserWithAccess+$FullAccessPermission
     if($FullAccessPermissions.indexof($FullAccessPermission) -lt (($FullAccessPermissions.count)-1))
     {
       $UserWithAccess=$UserWithAccess+","
     }
    }
    Print_Output
   }
  }

  #Getting delegated SendAs permission for mailbox
  if(($FilterPresent -eq 'False') -or ($SendAs.IsPresent))
  {
   $SendAsPermissions=(Get-RecipientPermission -Identity $upn | where{ -not (($_.Trustee -match "NT AUTHORITY") -or ($_.Trustee -match "S-1-5-21"))}).Trustee
   if([string]$SendAsPermissions -ne "")
   {
    $Print=1
    $UserWithAccess=""
    $AccessType="SendAs"
    foreach($SendAsPermission in $SendAsPermissions)
    {
     $UserWithAccess=$UserWithAccess+$SendAsPermission
     if($SendAsPermissions.indexof($SendAsPermission) -lt (($SendAsPermissions.count)-1))
     {
      $UserWithAccess=$UserWithAccess+","
     }
    }
    Print_Output
   }
  }

  #Getting delegated SendOnBehalf permission for mailbox
   if(($FilterPresent -eq 'False') -or ($SendOnBehalf.IsPresent))
   {
    $SendOnBehalfPermissions=$_.GrantSendOnBehalfTo
    if([string]$SendOnBehalfPermissions -ne "")
    {
     $Print=1
     $UserWithAccess=""
     $AccessType="SendOnBehalf"
     foreach($SendOnBehalfPermissionDN in $SendOnBehalfPermissions)
     {
      $SendOnBehalfPermission=(Get-Mailbox -Identity $SendOnBehalfPermissionDN).UserPrincipalName
      $UserWithAccess=$UserWithAccess+$SendOnBehalfPermission
      if($SendOnBehalfPermissions.indexof($SendOnBehalfPermission) -lt (($SendOnBehalfPermissions.count)-1))
      {
       $UserWithAccess=$UserWithAccess+","
      }
     }
     Print_Output
    }
   }
 }



function main{
 #Connect AzureAD and Exchange Online from PowerShell
 Get-PSSession | Remove-PSSession

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

 #Set output file
 $ExportCSV=".\MBPermission_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
 $Result="" 
 $Results=@()
 $MBUserCount=0
 $RolesAssigned=""

 #Check for AccessType filter
 if(($FullAccess.IsPresent) -or ($SendAs.IsPresent) -or ($SendOnBehalf.IsPresent))
 {}
 else
 {
  $FilterPresent='False'
 }

 #Check for input file
 if ($MBNamesFile -ne "") 
 { 
  #We have an input file, read it into memory 
  $MBs=@()
  $MBs=Import-Csv -Header "DisplayName" $MBNamesFile
  foreach($item in $MBs)
  {
   Get-Mailbox -Identity $item.displayname | Foreach{
   $MBUserCount++
   Get_MBPermission
   }
  }
 }
 #Getting all User mailbox
 else
 {
  Get-mailbox -ResultSize Unlimited | Where{$_.DisplayName -notlike "Discovery Search Mailbox"} | foreach{
   $MBUserCount++
   Get_MBPermission}
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
  Write-Host No mailbox found that matches your criteria.
}
#Clean up session 
Get-PSSession | Remove-PSSession
}
 . main
 