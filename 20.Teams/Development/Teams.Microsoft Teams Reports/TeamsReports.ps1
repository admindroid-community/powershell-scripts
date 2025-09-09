<#
=============================================================================================
Name:           Export Microsoft Teams report
Description:    This script exports Microsoft Teams report to CSV
Version:        2.0
website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~

1.A single script allows you to generate eight different Teams reports.  
2.The script can be executed with MFA enabled accounts too. 
3.Exports output to CSV. 
4.Automatically installs Microsoft Teams PowerShell module (if not installed already) upon your confirmation. 
5.The script is scheduler friendly. I.e., Credential can be passed as a parameter instead of saving inside the script.
6.The script supports certificate-based authentication. 


For detailed Script execution: https://o365reports.com/2020/05/28/microsoft-teams-reporting-using-powershell/
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

#Connect to Microsoft Teams
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
Write-Host Importing Microsoft Teams module... -ForegroundColor Yellow


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
If($Team -ne $null)
{
 Write-host `nSuccessfully connected to Microsoft Teams -ForegroundColor Green
}
else
{
 Write-Host Error occurred while creating Teams session. Please try again -ForegroundColor Red
 exit
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
 Write-host `nMicrosoft Teams Reporting -ForegroundColor Yellow
 Write-Host  "    1.All Teams in organization" -ForegroundColor Cyan
 Write-Host  "    2.All Teams members and owners report" -ForegroundColor Cyan
 Write-Host  "    3.Specific Teams' members and Owners report" -ForegroundColor Cyan
 Write-Host  "    4.All Teams' owners report" -ForegroundColor Cyan
 Write-Host  "    5.Specific Teams' owners report" -ForegroundColor Cyan
 Write-Host `nTeams Channel Reporting -ForegroundColor Yellow
 Write-Host  "    6.All channels in an organization" -ForegroundColor Cyan
 Write-Host  "    7.All channels in a specific Team" -ForegroundColor Cyan
 Write-Host  "    8.Members and Owners Report of a Single Channel" -ForegroundColor Cyan
 Write-Host  "    0.Exit" -ForegroundColor Cyan
 Write-Host `nPrivate Channel Management and Reporting -ForegroundColor Yellow
 Write-Host  "    You can download the script from https://blog.admindroid.com/managing-private-channels-in-microsoft-teams/" -ForegroundColor Cyan

 $i = Read-Host `n'Please choose the action to continue' 
 }
 else
 {
  $i=$Action
 }

 Switch ($i) {
  1 {
     $Result=""  
     $Results=@() 
     $Path="./All Teams Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Write-Host Exporting all Teams report...
     $Count=0
     Get-Team | foreach {
     $TeamName=$_.DisplayName
     Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName"
     $Count++
     $Visibility=$_.Visibility
     $MailNickName=$_.MailNickName
     $Description=$_.Description
     $Archived=$_.Archived
     $GroupId=$_.GroupId
     $ChannelCount=(Get-TeamChannel -GroupId $GroupId).count
     $TeamUser=Get-TeamUser -GroupId $GroupId
     $TeamMemberCount=$TeamUser.Count
     $TeamOwnerCount=($TeamUser | ?{$_.role -eq "Owner"}).count
     $Result=@{'Teams Name'=$TeamName;'Team Type'=$Visibility;'Mail Nick Name'=$MailNickName;'Description'=$Description;'Archived Status'=$Archived;'Channel Count'=$ChannelCount;'Team Members Count'=$TeamMemberCount;'Team Owners Count'=$TeamOwnerCount}
     $Results= New-Object psobject -Property $Result
     $Results | select 'Teams Name','Team Type','Mail Nick Name','Description','Archived Status','Channel Count','Team Members Count','Team Owners Count' | Export-Csv $Path -NoTypeInformation -Append
     }
     Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName" -Completed
     if((Test-Path -Path $Path) -eq "True") 
     {
      Write-Host `nReport available in $Path -ForegroundColor Green
     }
    }
  2 {
     $Result=""  
     $Results=@() 
     $Path="./All Teams Members and Owner Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Write-Host Exporting all Teams members and owners report...
     $Count=0
     Get-Team | foreach {
      $TeamName=$_.DisplayName
      Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName"
      $Count++
      $GroupId=$_.GroupId
      Get-TeamUser -GroupId $GroupId | foreach {
       $Name=$_.Name
       $MemberMail=$_.User
       $Role=$_.Role
       $Result=@{'Teams Name'=$TeamName;'Member Name'=$Name;'Member Mail'=$MemberMail;'Role'=$Role}
       $Results= New-Object psobject -Property $Result
       $Results | select 'Teams Name','Member Name','Member Mail','Role' | Export-Csv $Path -NoTypeInformation -Append
      }
     }
     Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName" -Completed
     if((Test-Path -Path $Path) -eq "True") 
     {
      Write-Host `nReport available in $Path -ForegroundColor Green
     }
    }

  3 {
     $Result=""  
     $Results=@() 
     $TeamName=Read-Host Enter Teams name to get members report "(Case sensitive)":
     $GroupId=(Get-Team -DisplayName $TeamName).GroupId 
     Write-Host Exporting $TeamName"'s" Members and Owners report...
     $Path=".\MembersOf $TeamName Team Report _$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-TeamUser -GroupId $GroupId | foreach {
      $Name=$_.Name
      $MemberMail=$_.User
      $Role=$_.Role
      $Result=@{'Member Name'=$Name;'Member Mail'=$MemberMail;'Role'=$Role}
      $Results= New-Object psobject -Property $Result
      $Results | select 'Member Name','Member Mail','Role' | Export-Csv $Path -NoTypeInformation -Append
     }
     if((Test-Path -Path $Path) -eq "True") 
     {
      Write-Host `nReport available in $Path -ForegroundColor Green
     }
    }

  4 {
     $Result=""  
     $Results=@() 
     $Path="./All Teams Owner Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Write-Host Exporting all Teams owner report...
     $Count=0
     Get-Team | foreach {
      $TeamName=$_.DisplayName
      Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName"
      $Count++
      $GroupId=$_.GroupId
      Get-TeamUser -GroupId $GroupId | ?{$_.role -eq "Owner"} | foreach {
       $Name=$_.Name
       $MemberMail=$_.User
       $Result=@{'Teams Name'=$TeamName;'Owner Name'=$Name;'Owner Mail'=$MemberMail}
       $Results= New-Object psobject -Property $Result
       $Results | select 'Teams Name','Owner Name','Owner Mail' | Export-Csv $Path -NoTypeInformation -Append
      }
     }
     Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName" -Completed
     if((Test-Path -Path $Path) -eq "True") 
     {
      Write-Host `nReport available in $Path -ForegroundColor Green
     }
    }

  5 {
     $Result=""  
     $Results=@() 
     $TeamName=Read-Host Enter Teams name to get owners report "(Case sensitive)":
     $GroupId=(Get-Team -DisplayName $TeamName).GroupId 
     Write-Host Exporting $TeamName team"'"s Owners report...
     $Path=".\OwnersOf $TeamName team report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-TeamUser -GroupId $GroupId | ?{$_.role -eq "Owner"} | foreach {
      $Name=$_.Name
      $MemberMail=$_.User
      $Result=@{'Member Name'=$Name;'Member Mail'=$MemberMail}
      $Results= New-Object psobject -Property $Result
      $Results | select 'Member Name','Member Mail' | Export-Csv $Path -NoTypeInformation -Append
     }
     if((Test-Path -Path $Path) -eq "True") 
     {
      Write-Host `nReport available in $Path -ForegroundColor Green
     }
    }

  6 {
      $Result=""  
      $Results=@() 
      $Path="./All Channels Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Write-Host Exporting all Channels report...
      $Count=0
      Get-Team | foreach {
       $TeamName=$_.DisplayName
       Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing Team: $TeamName "
       $Count++
       $GroupId=$_.GroupId
       Get-TeamChannel -GroupId $GroupId | foreach {
        $ChannelName=$_.DisplayName
        Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing Team: $TeamName "`n" Currently Processing Channel: $ChannelName"
        $MembershipType=$_.MembershipType
        $Description=$_.Description
        $ChannelUser=Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName
        $ChannelMemberCount=$ChannelUser.Count
        $ChannelOwnerCount=($ChannelUser | ?{$_.role -eq "Owner"}).count
        $Result=@{'Teams Name'=$TeamName;'Channel Name'=$ChannelName;'Membership Type'=$MembershipType;'Description'=$Description;'Total Members Count'=$ChannelMemberCount;'Owners Count'=$ChannelOwnerCount}
        $Results= New-Object psobject -Property $Result
        $Results | select 'Teams Name','Channel Name','Membership Type','Description','Owners Count','Total Members Count' | Export-Csv $Path -NoTypeInformation -Append
       }
      }
      Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName  `n Currently Processing Channel: $ChannelName"  -Completed
      if((Test-Path -Path $Path) -eq "True") 
      {
       Write-Host `nReport available in $Path -ForegroundColor Green
      }
     }  

   7 {
      $TeamName=Read-Host Enter Teams name "(Case Sensitive)"
      Write-Host Exporting Channels report...
      $Count=0
      $GroupId=(Get-Team -DisplayName $TeamName).GroupId
      $Path=".\Channels available in $TeamName team $((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Get-TeamChannel -GroupId $GroupId | Foreach {
       $ChannelName=$_.DisplayName
       Write-Progress -Activity "`n     Processed channel count: $Count "`n"  Currently Processing Channel: $ChannelName"
       $Count++
       $MembershipType=$_.MembershipType
       $Description=$_.Description
       $ChannelUser=Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName
       $ChannelMemberCount=$ChannelUser.Count
       $ChannelOwnerCount=($ChannelUser | ?{$_.role -eq "Owner"}).count
       $Result=@{'Teams Name'=$TeamName;'Channel Name'=$ChannelName;'Membership Type'=$MembershipType;'Description'=$Description;'Total Members Count'=$ChannelMemberCount;'Owners Count'=$ChannelOwnerCount}
       $Results= New-Object psobject -Property $Result
       $Results | select 'Teams Name','Channel Name','Membership Type','Description','Owners Count','Total Members Count' | Export-Csv $Path -NoTypeInformation -Append
      }
      Write-Progress -Activity "`n     Processed channel count: $Count "`n"  Currently Processing Channel: $ChannelName" -Completed
      if((Test-Path -Path $Path) -eq "True") 
      {
       Write-Host `nReport available in $Path -ForegroundColor Green
      }
     }  
    
   8 {
      $Result=""  
      $Results=@() 
      $TeamName=Read-Host Enter Teams name in which Channel resides "(Case sensitive)"
      $ChannelName=Read-Host Enter Channel name
      $GroupId=(Get-Team -DisplayName $TeamName).GroupId 
      Write-Host Exporting $ChannelName"'s" Members and Owners report...
      $Path=".\MembersOf $ChannelName channel report $((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName | foreach {
       $Name=$_.Name
       $UPN=$_.User
       $Role=$_.Role
       $Result=@{'Teams Name'=$TeamName;'Channel Name'=$ChannelName;'Member Mail'=$UPN;'Member Name'=$Name;'Role'=$Role}
       $Results= New-Object psobject -Property $Result
       $Results | select 'Teams Name','Channel Name','Member Name','Member Mail',Role | Export-Csv $Path -NoTypeInformation -Append
      }   
      if((Test-Path -Path $Path) -eq "True") 
      {
	   Write-Host ""
       Write-Host "Report available in:"  -NoNewline -ForegroundColor Yellow
	   Write-Host $Path 
	   Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n 
      }
     }

   }
   if($Action -ne "")
   {exit}
}
  While ($i -ne 0)
  Clear-Host
 