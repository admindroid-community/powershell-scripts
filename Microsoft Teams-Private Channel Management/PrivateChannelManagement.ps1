<#
=============================================================================================
Name:           Private Channel Management & Reporting Using PowerShell
Description:    This script performs Private Channel related management actions and reporting
Version:        2.0
Script by:      AdminDroid Team 


To run the script
./PrivateChannelManagement.ps1

To schdeule/run the script by explicitly mentioning credential
./PrivateChannelManagement.ps1 -UserName <UserName> -Password <Password>

To run the script with certificate based authentication
./PrivateChannelManagement.ps1 -TenantId <TenantId> -AppId <AppId> -CertificateThumbPrint <CertThumbPrint>

To run a specific action directly
./PrivateChannelManagement.ps1 -Action 7


Change Log
~~~~~~~~~~
 V1.0 (Nov 18, 2019) - File created
 V2.0 (Nov 13, 2024) - Added support for certificate-based authentication, removed older PowerShell modules, and a few minor usability enhancements


For detailed script execution steps: https://blog.admindroid.com/managing-private-channels-in-microsoft-teams/
============================================================================================
#>

#Accept input paramenters 
param(
[string]$UserName, 
[string]$Password, 
[string]$TenantId,
[string]$AppId,
[string]$CertificateThumbprint,
[int]$Action
) 


Function MSTeam_PSModule
{
 #Check for MS Teams PowerShell module availability
 $Module=Get-Module -Name MicrosoftTeams -ListAvailable 
 if($Module.count -eq 0)
 {
  Write-Host MicrosoftTeams module is not available  -ForegroundColor yellow 
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
  if($Confirm -match "[yY]")
  {
   Install-Module MicrosoftTeams -Scope CurrentUser
  }
  else
  {
   Write-Host MicrosoftTeams module is required.Please install module using Install-Module MicrosoftTeams cmdlet.
   Exit
  }
 }
 Write-Host Connecting to Microsoft Teams... -ForegroundColor Yellow


 #Authentication using non-MFA

 #Storing credential in script for scheduling purpose/ Passing credential as parameter
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  $Team=Connect-MicrosoftTeams -Credential $Credential
 }
 elseif(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  $Team=Connect-MicrosoftTeams  -TenantId $TenantId -ApplicationId $AppId -CertificateThumbprint $CertificateThumbprint 
 }
 else
 {  
  $Team=Connect-MicrosoftTeams
 }


 #Check for Teams connectivity
 If($Team -eq $null)
 {
  Write-Host Error occurred while creating Teams session. Please try again -ForegroundColor Red
  exit
 }
}

Function Open_Output
{
  if((Test-Path -Path $Path) -eq "True") 
 {
  Write-Host `nThe exported report available in: -NoNewline -Foregroundcolor Yellow; Write-Host $ExportCSV

  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 
   $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$Path"   
  } 
 }
 else
 {
  Write-Host No data found.
 }
}

MSTeam_PSModule
$Location=Get-Location
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
     Write-Host `nThe file must follow the format: Users"'" UPN separated by new line -ForegroundColor Magenta
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
      $Path="$Location/PrivateChannelsReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Write-Host Exporting Private Channels report...
      $Count=0
      Get-Team | foreach {
      $TeamName=$_.DisplayName
      $Count++
      Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName"
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
     Open_Output
    }  

   8 {
      $TeamName=Read-Host Enter Teams name "(Case Sensitive)":
      Write-Host Exporting Private Channel report...
      $GroupId=(Get-Team -DisplayName $TeamName).GroupId
      $Path="$Location\Private Channels available in $TeamName$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Get-TeamChannel -GroupId $GroupId -MembershipType Private | select DisplayName | Export-Csv $Path -NoTypeInformation
       Open_Output
     }

   9{
     $Result=""  
     $Results=@() 
     Write-Host Exporting all Private Channel"'s" Members and Owners report...
     $Count=0
     $Path="$Location/AllPrivateChannels Members and Owners Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-Team | foreach {
      $TeamName=$_.DisplayName
      $GroupId=$_.GroupId
      $PrivateChannels=(Get-TeamChannel -GroupId $GroupId -MembershipType Private).DisplayName
      foreach($PrivateChannel in $PrivateChannels)
      {
       $Count++
       Write-Progress -Activity "`n     Processed Private Channel count: $Count "`n"  Currently Processing: $PrivateChannel"
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
   Open_Output
    }    

   10 {
    $Result=""  
    $Results=@() 
    $TeamName=Read-Host Enter Teams name in which Private Channel resides "(Case sensitive)":
    $ChannelName=Read-Host Enter Private Channel name
    $GroupId=(Get-Team -DisplayName $TeamName).GroupId 
    Write-Host Exporting $ChannelName"'s" Members and Owners report...
    $Path="$Location\MembersOf $ChannelName$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
    Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName | foreach {
     $Name=$_.Name
     $UPN=$_.User
     $Role=$_.Role
     $Result=@{'Teams Name'=$TeamName;'Private Channel Name'=$ChannelName;'UPN'=$UPN;'User Display Name'=$Name;'Role'=$Role}
     $Results= New-Object psobject -Property $Result
     $Results | select 'Teams Name','Private Channel Name',UPN,'User Display Name',Role | Export-Csv $Path -NoTypeInformation -Append
    }   
     Open_Output
   }
  }
  if($Action -ne "")
  {exit}
 }
  While ($i -ne 0)


 

