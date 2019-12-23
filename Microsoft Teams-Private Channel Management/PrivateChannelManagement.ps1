<#
=============================================================================================
Name:           PrivateChannelManagement
Description:    This script performs Private Channel related management actions and reporting
Version:        1.0
Released date:  18/11/2019
website:        blog.admindroid.com
Script by:      AdminDroid Team (Proud Creators of AdminDroid Office 365 Reporting Tool)


To run the script
./PrivateChannelManagement.ps1

To schdeule/run the script by explicitly mentioning credential
./PrivateChannelManagement.ps1 -UserName <UserName> -Password <Password>

To run the script with MFA enabled account
./PrivateChannelManagement.ps1 -MFA
============================================================================================
#>

#Accept input paramenters 
param(
[string]$UserName, 
[string]$Password, 
[switch]$MFA,
[int]$Action
) 


#install latest Microsoft Teams module from PowerShell Test Gallery
$Module=Get-Module -Name MicrosoftTeams -ListAvailable  
if($Module.count -eq 0) 
{   
 $Confirm= Read-Host Are you sure you want to install Microsoft Teams module? [Y] Yes [N] No 
 if($Confirm -match "[y]") 
 { 
  Register-PSRepository -Name PSGalleryInt -SourceLocation https://www.poshtestgallery.com/ -InstallationPolicy Trusted
  Install-Module -Name MicrosoftTeams -Repository PSGalleryInt -Force
  Write-Host Installing Microsoft Teams Module...
 }
 else
 {
  Write-Host `nNeed Microsoft Teams PowerShell module. Please install the latest module from PowerShell Test Gallery -ForegroundColor Yellow
  exit
 }
}
#Check for latest Microsoft Teams PowerShell Module
elseif((Get-module -Name MicrosoftTeams -ListAvailable).version -lt "1.0.18")
{
 Write-Host `nTo manage Private Channel, you must install lastest version of MicrosoftTeams PowerShell module from PowerShell Test Gallery
 $Confirm= Read-Host `nAre you sure you want to uninstall old version ? [Y] Yes [N] No 
 if($Confirm -match "[y]") 
 { 
  Uninstall-Module -Name MicrosoftTeams
 }
 else
 {
  Write-Host Please install latest version of Microsoft Teams PowerShell module from PowerShell Test Gallery -ForegroundColor Yellow
  exit
 }
 $Confirm= Read-Host `nAre you sure you want to install latest version of Microsoft Teams module? [Y] Yes [N] No 
 if($Confirm -match "[y]") 
 { 
  Register-PSRepository -Name PSGalleryInt -SourceLocation https://www.poshtestgallery.com/ -InstallationPolicy Trusted
  Install-Module -Name MicrosoftTeams -Repository PSGalleryInt -Force
  Write-Host Installing Microsoft Teams Module...
 }
 else
 {
  Write-Host `nPlease install latest version of Microsoft Teams PowerShell module from PowerShell Test Gallery -ForegroundColor Yellow
  exit
 }
}


#Check Skype for Business Online module
$Module=Get-Module -Name SkypeOnlineConnector -ListAvailable 
if($Module.count -eq 0) 
{ 
 Write-Host  `nPlease install Skype for Business Online PowerShell Module  -ForegroundColor yellow  
 Write-Host `nYou can download the Skype Online PowerShell module directly using below url: https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowerShell.Exe 
 Write-Host `nAfter installing module, Please close all existing PowerShell sessions. Start new PowerShell console and rerun this script.
 exit
} 
Write-Host Preparing required PowerShell Modules...
Get-PSSession | Remove-PSSession
#Authentication using MFA
if($MFA.IsPresent) 
{ 
 Write-Host Importing Skype for Business Online PowerShell Module...
 $sfbSession = New-CsOnlineSession 
 Import-PSSession $sfbSession -AllowClobber | Out-Null 
 Write-Host Importing Microsoft Teams PowerShell Module...
 Connect-MicrosoftTeams | Out-Null
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
 Write-Host Importing Skype for Business Online PowerShell Module...
 $sfbSession = New-CsOnlineSession -Credential $Credential 
 Import-PSSession $sfbSession -AllowClobber -WarningAction SilentlyContinue | Out-Null 
 Write-Host Importing Microsoft Teams PowerShell Module...
 Connect-MicrosoftTeams -Credential $Credential | Out-Null
} 
[boolean]$Delay=$false
Do {
 if($Action -eq "")
 {
  if($Delay -eq $true)
  {
   Start-Sleep -Seconds 2
  }
  $Delay=$true
 Write-Host ""
 Write-host `nPrivate Channel Management -ForegroundColor Yellow
 Write-Host  "    1.Allow for Organization" -ForegroundColor Cyan
 Write-Host  "    2.Disable for Organization" -ForegroundColor Cyan
 Write-Host  "    3.Allow for a User" -ForegroundColor Cyan
 Write-Host  "    4.Disable for a User" -ForegroundColor Cyan
 Write-Host  "    5.Allow User in bulk using CSV import" -ForegroundColor Cyan
 Write-Host  "    6.Disable User in bulk using CSV import" -ForegroundColor Cyan
 Write-Host `nPrivate Channel Reporting -ForegroundColor Yellow
 Write-Host  "    7.All Private Channels in Organization" -ForegroundColor Cyan
 Write-Host  "    8.All Private Channels in Teams" -ForegroundColor Cyan
 Write-Host  "    9.Members and Owners Report of All Private Channels" -ForegroundColor Cyan
 Write-Host  "    10.Members and Owners Report of Single Private Channel" -ForegroundColor Cyan
 Write-Host  "    0.Exit" -ForegroundColor Cyan
 Write-Host ""
 $i = Read-Host 'Please choose the action to continue' 
 }
 else
 {
  $i=$Action
 }
 Switch ($i) {
  1 {
     Set-CsTeamsChannelsPolicy -Identity Global –AllowPrivateChannelCreation $True 
     Write-Host Private Channel creation allowed for Organization wide -ForegroundColor Green    
    }
      
  2 {
       Set-CsTeamsChannelsPolicy -Identity Global –AllowPrivateChannelCreation $False 
       Write-Host Blocked Private Channel creation Organization wide -ForegroundColor Green 
    }

  3 {
     if((Get-CsTeamsChannelsPolicy -Identity "Allow Private Channel Creation" -ErrorAction silent) -eq $null)
     {
      New-CsTeamsChannelsPolicy -Identity "Allow Private Channel Creation" -AllowPrivateChannelCreation $True | Out-Null
     }  
     $User=Read-Host Enter User name"(UPN format)" to grant the right to create Private Channel
     Grant-CsTeamsChannelsPolicy -PolicyName "Allow Private Channel Creation" -Identity $User 
     if($?)
     {
      Write-host `nNow $User can create Private Channel -ForegroundColor Green
     }  
    }

  4 {
     if((Get-CsTeamsChannelsPolicy -Identity "Disable Private Channel Creation" -ErrorAction silent) -eq $null)
     {
      New-CsTeamsChannelsPolicy -Identity "Disable Private Channel Creation" -AllowPrivateChannelCreation $False | Out-Null
     }  
     $User=Read-Host Enter User name"(UPN format)" to disable Private Channel creation
     Grant-CsTeamsChannelsPolicy -PolicyName "Disable Private Channel Creation" -Identity $User
     if($?)
     {
      $?
      Write-host `nPrivate Channel creation blocked for $User -ForegroundColor green
     }
    }

  5 {
     if((Get-CsTeamsChannelsPolicy -Identity "Allow Private Channel Creation" -ErrorAction silent) -eq $null)
     {
      New-CsTeamsChannelsPolicy -Identity "Allow Private Channel Creation" -AllowPrivateChannelCreation $True
     }
     Write-Host `nThe file must follow the format: Users"'" UPN separated by new line without header -ForegroundColor Magenta
     $UserNamesFile=Read-Host Enter CSV/txt file path"(Eg:C:\Users\Desktop\UserNames.txt)"
     $Users=@()
     $Users=Import-Csv -Header "UserPrincipalName" $UserNamesFile
     foreach($User in $Users)
     {
      Grant-CsTeamsChannelsPolicy -PolicyName "Allow Private Channel Creation" -Identity $User.UserPrincipalName
      if($?)
      {
       Write-host Now $User.UserPrincipalName can create Private Channel -ForegroundColor Green
      }
     }
    }   

  6 {
     if((Get-CsTeamsChannelsPolicy -Identity "Disable Private Channel Creation" -ErrorAction silent) -eq $null)
     {
      New-CsTeamsChannelsPolicy -Identity "Disable Private Channel Creation" -AllowPrivateChannelCreation $False
     }
     Write-Host `nThe file must follow the format: Users"'" UPN separated by new line without header -ForegroundColor Magenta
     $UserNamesFile=Read-Host Enter CSV/txt file path"(Eg:C:\Users\Desktop\UserNames.txt)"
     $Users=@()
     $Users=Import-Csv -Header "UserPrincipalName" $UserNamesFile
     foreach($User in $Users)
     {
      Grant-CsTeamsChannelsPolicy -PolicyName "Disable Private Channel Creation" -Identity $User.UserPrincipalName
      if($?)
      {
       Write-host Private Channel creation blocked for $User.UserPrincipalName -ForegroundColor Green
      }
     } 
    }
       
   7 {
      $Result=""  
      $Results=@() 
      $Path="./AllPrivateChannels_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Write-Host Exporting Private Channels report...
      $Count=0
      Get-Team | foreach {
      $TeamName=$_.DisplayName
      Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName"
      $Count++
      $GroupId=$_.GroupId
      $PrivateChannels=(Get-TeamChannel -GroupId $GroupId -MembershipType Private).DisplayName
      foreach($PrivateChannel in $PrivateChannels)
      {
       $Result=@{'Teams Name'=$TeamName;'Private Channel'=$PrivateChannel}
       $Results= New-Object psobject -Property $Result
       $Results | select 'Teams Name','Private Channel' | Export-Csv $Path -NoTypeInformation -Append
      }
     }
     Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName" -Completed
     if((Test-Path -Path $Path) -eq "True") 
     {
      Write-Host `nReport available in $Path -ForegroundColor Green
     }
    }  

   8 {
      $TeamName=Read-Host Enter Teams name "(Case Sensitive)":
      Write-Host Exporting Private Channel report...
      $GroupId=(Get-Team -DisplayName $TeamName).GroupId
      $Path=".\Private Channels available in $TeamName$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Get-TeamChannel -GroupId $GroupId -MembershipType Private | select DisplayName | Export-Csv $Path -NoTypeInformation
      if((Test-Path -Path $Path) -eq "True") 
      {
       Write-Host `nReport available in $Path -ForegroundColor Green
      }
     }

   9{
     $Result=""  
     $Results=@() 
     Write-Host Exporting all Private Channel"'s" Members and Owners report...
     $Count=0
     $Path="./AllPrivateChannels Members and Owners Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-Team | foreach {
      $TeamName=$_.DisplayName
      $GroupId=$_.GroupId
      $PrivateChannels=(Get-TeamChannel -GroupId $GroupId -MembershipType Private).DisplayName
      foreach($PrivateChannel in $PrivateChannels)
      {
       Write-Progress -Activity "`n     Processed Private Channel count: $Count "`n"  Currently Processing: $PrivateChannel"
       $Count++
       Get-TeamChannelUser -GroupId $GroupId -DisplayName $PrivateChannel | foreach {
        $Name=$_.Name
        $UPN=$_.User
        $Role=$_.Role
        $Result=@{'Teams Name'=$TeamName;'Private Channel Name'=$PrivateChannel;'UPN'=$UPN;'User Display Name'=$Name;'Role'=$Role}
        $Results= New-Object psobject -Property $Result
        $Results | select 'Teams Name','Private Channel Name',UPN,'User Display Name',Role | Export-Csv $Path -NoTypeInformation -Append
       }
      }    
     }
     Write-Progress -Activity "`n     Processed Private Channel count: $Count "`n"  Currently Processing: $PrivateChannel" -Completed
     if((Test-Path -Path $Path) -eq "True") 
     {
      Write-Host `nReport available in $Path -ForegroundColor Green
     }
    }    

   10 {
    $Result=""  
    $Results=@() 
    $TeamName=Read-Host Enter Teams name in which Private Channel resides "(Case sensitive)":
    $ChannelName=Read-Host Enter Private Channel name
    $GroupId=(Get-Team -DisplayName $TeamName).GroupId 
    Write-Host Exporting $ChannelName"'s" Members and Owners report...
    $Path=".\MembersOf $ChannelName$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
    Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName | foreach {
     $Name=$_.Name
     $UPN=$_.User
     $Role=$_.Role
     $Result=@{'Teams Name'=$TeamName;'Private Channel Name'=$ChannelName;'UPN'=$UPN;'User Display Name'=$Name;'Role'=$Role}
     $Results= New-Object psobject -Property $Result
     $Results | select 'Teams Name','Private Channel Name',UPN,'User Display Name',Role | Export-Csv $Path -NoTypeInformation -Append
    }   
    if((Test-Path -Path $Path) -eq "True") 
    {
     Write-Host `nReport available in $Path -ForegroundColor Green
    }
   }
  }
  if($Action -ne "")
  {exit}
 }
  While ($i -ne 0)
  Clear-Host
 
