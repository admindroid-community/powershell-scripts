<#
=============================================================================================
Name:           Export MS Teams Private Channels and Membership Report
Description:    This script exports 7 private channel reports
Version:        1.0
website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~

1.A single script allows you to generate 7 different Private channels reports.  
2.The script can be executed with MFA enabled accounts too. 
3.Exports output to CSV. 
4.Automatically installs Microsoft Teams PowerShell module (if not installed already) upon your confirmation. 
5.The script is scheduler friendly. i.e., Credential can be passed as a parameter instead of saving inside the script.
6.The script supports certificate-based authentication. 



For detailed Script execution: https://o365reports.com/2024/01/09/export-microsoft-teams-private-channel-membership-reports-using-powershell
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

Function CheckOutput
{
 if((Test-Path -Path $Path) -eq "True")
 {
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6)
    {
        Invoke-Item "$Path"
    }
    Write-Host "Detailed report available in: $Path" -ForegroundColor Green
 }
 else
 {
    Write-Host "No data found" -ForegroundColor Red
 }
}


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
Import-Module MicrosoftTeams


 #Storing credential in script for scheduling purpose/ Passing credential as parameter
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  $Team=Connect-MicrosoftTeams -Credential $Credential
 }
 #Authentication using CBA
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
 Write-host `nMicrosoft Teams Private Channel Reporting -ForegroundColor Yellow
 Write-Host  "    1.Export all Private Channels" -ForegroundColor Cyan
 Write-Host  "    2.Export Private Channels in a specific team" -ForegroundColor Cyan
 Write-Host  "    3.Export all Private Channels & their membership report" -ForegroundColor Cyan
 Write-Host  "    4.Export membership of Private Channels in a Specific team" -ForegroundColor Cyan
 Write-Host  "    5.Export Private Channels' owners report" -ForegroundColor Cyan
 Write-Host  "    6.Export Private Channels' with guests" -ForegroundColor Cyan
 Write-Host  "    7.Export all teams with Private Channels" -ForegroundColor Cyan
 Write-Host  "    0.Exit" -ForegroundColor Cyan

 $i = Read-Host `n'Please choose the action to continue' 
 }
 else
 {
  $i=$Action
 }
 $Location=Get-Location
 Switch ($i) {
 1 {
      $Result=""  
      $Results=@() 
      $Path="$Location\All Private Channels Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Write-Host Exporting private Channels report...
      $Count=0
      Get-Team | foreach {
       $TeamName=$_.DisplayName
       Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing Team: $TeamName "
       $Count++
       $GroupId=$_.GroupId
       Get-TeamChannel -GroupId $GroupId -MembershipType Private | foreach {
        $ChannelName=$_.DisplayName
        Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing Team: $TeamName "`n" Currently Processing Channel: $ChannelName"
        $MembershipType=$_.MembershipType
        $Description=$_.Description
        $ChannelUser=Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName
        $ChannelMemberCount=$ChannelUser.Count
        $ChannelOwnerCount=($ChannelUser | ?{$_.role -eq "Owner"}).count
        $ChannelGuestCount=($ChannelUser | ?{$_.role -eq "Guest"}).count
        $Result=@{'Teams Name'=$TeamName;'Channel Name'=$ChannelName;'Membership Type'=$MembershipType;'Description'=$Description;'Total Members Count'=$ChannelMemberCount;'Owners Count'=$ChannelOwnerCount;'Guests Count'=$ChannelGuestCount}
        $Results= New-Object psobject -Property $Result
        $Results | select 'Teams Name','Channel Name','Membership Type','Description','Owners Count','Guests Count','Total Members Count' | Export-Csv $Path -NoTypeInformation -Append
       }
      }
      Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName  `n Currently Processing Channel: $ChannelName"  -Completed
      CheckOutput
     }  

   2 {
      $TeamName=Read-Host Enter Teams name "(Case Sensitive)"
      Write-Host Exporting private channels...
      $Count=0
      $GroupId=(Get-Team -DisplayName $TeamName).GroupId
      $Path="$Location\PrivateChannels available in $TeamName team $((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Get-TeamChannel -GroupId $GroupId -MembershipType Private | Foreach {
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
      CheckOutput
     }  
  3 {
     $Result=""  
     $Results=@() 
     Write-Host Exporting all Teams members and owners report...
     $Count=0
     $Path="$Location\PrivateChannels Membership Report $((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-Team | foreach {
      $TeamName=$_.DisplayName
      Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName"
      $Count++
      $GroupId=$_.GroupId
      Get-TeamChannel -GroupId $GroupId -MembershipType Private | Foreach {
       $ChannelName=$_.DisplayName
       Write-Progress -Activity "`n Processed channel count: $Count    "`n"  Currently Processing Channel: $ChannelName"
       $Count++
       $MembershipType=$_.MembershipType
       $Description=$_.Description
       Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName | foreach{
        $Name=$_.Name
        $MemberMail=$_.User
        $Role=$_.Role
        $Result=@{'Teams Name'=$TeamName;'Channel Name'=$ChannelName;'Member Name'=$Name;'Member Mail'=$MemberMail;'Role'=$Role}
        $Results= New-Object psobject -Property $Result
        $Results | select 'Teams Name','Channel Name','Member Name','Member Mail','Role' | Export-Csv $Path -NoTypeInformation -Append
       }
      }
     }
     Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName" -Completed
     CheckOutput
    }
  4 {
     $Result=""  
     $Results=@() 
     Write-Host Exporting membership of Private Channels in a specific team...
     $Count=0
     $TeamName=Read-Host Enter Teams name in which Channel resides "(Case sensitive)"
     $GroupId=(Get-Team -DisplayName $TeamName).GroupId 
     $Path="$Location\Membership Report of Private Channels in a $TeamName team $((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-TeamChannel -GroupId $GroupId -MembershipType Private | Foreach {
      $Count++
      $ChannelName=$_.DisplayName
      Write-Progress -Activity "`n     Processed channel count: $Count "`n"  Currently Processing Channel: $ChannelName"
      $MembershipType=$_.MembershipType
      $Description=$_.Description
      Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName | foreach{
       $Name=$_.Name
       $MemberMail=$_.User
       $Role=$_.Role
       $Result=@{'Teams Name'=$TeamName;'Channel Name'=$ChannelName;'Member Name'=$Name;'Member Mail'=$MemberMail;'Role'=$Role}
       $Results= New-Object psobject -Property $Result
       $Results | select 'Channel Name','Member Name','Member Mail','Role','Teams Name' | Export-Csv $Path -NoTypeInformation -Append
      }
     }
    Write-Progress -Activity "`n     Processed channel count: $Count "`n"  Currently Processing Channel: $ChannelName" -Completed
     CheckOutput
    }

  5 {
     $Result=""  
     $Results=@() 
     $Path="$Location\Privatec Channels Owner Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Write-Host "Exporting private channels' owner report..."
     $Count=0
     $Path=".\PrivateChannels Ownership Report $((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-Team | foreach {
      $TeamName=$_.DisplayName
      Write-Progress -Activity "`n  Processed Teams count: $Count "`n"  Currently Processing: $TeamName"
      $Count++
      $GroupId=$_.GroupId
      Get-TeamChannel -GroupId $GroupId -MembershipType Private | Foreach {
       $ChannelName=$_.DisplayName
       Write-Progress -Activity "`n  Processed Teams count: $Count   "`n"  Currently Processing Channel: $ChannelName"
       $Count++
       $MembershipType=$_.MembershipType
       $Description=$_.Description
       Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -Role Owner | foreach{
        $Name=$_.Name
        $MemberMail=$_.User
        $Role=$_.Role
        $Result=@{'Teams Name'=$TeamName;'Channel Name'=$ChannelName;'Member Name'=$Name;'Member Mail'=$MemberMail;'Role'=$Role}
        $Results= New-Object psobject -Property $Result
        $Results | select 'Teams Name','Channel Name','Member Name','Member Mail','Role' | Export-Csv $Path -NoTypeInformation -Append
       }
      }
     }
     Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName" -Completed
     CheckOutput
    }

    6 {
     $Result=""  
     $Results=@() 
     $Path="$Location\Privatec Channels Owner Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Write-Host "Exporting private channels' with guest users..."
     $Count=0
     $Path=".\PrivateChannels with Guests Report $((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-Team | foreach {
      $TeamName=$_.DisplayName
      Write-Progress -Activity "`n  Processed Teams count: $Count "`n"  Currently Processing: $TeamName"
      $Count++
      $GroupId=$_.GroupId
      Get-TeamChannel -GroupId $GroupId -MembershipType Private | Foreach {
       $ChannelName=$_.DisplayName
       Write-Progress -Activity "`n  Processed Teams count: $Count   "`n"  Currently Processing Channel: $ChannelName"
       $Count++
       $MembershipType=$_.MembershipType
       $Description=$_.Description
       Get-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -Role Guest | foreach{
        $Name=$_.Name
        $MemberMail=$_.User
        $Role=$_.Role
        $Result=@{'Teams Name'=$TeamName;'Channel Name'=$ChannelName;'Guest Name'=$Name;'Guest Mail'=$MemberMail;'Role'=$Role}
        $Results= New-Object psobject -Property $Result
        $Results | select 'Teams Name','Channel Name','Guest Name','Guest Mail','Role' | Export-Csv $Path -NoTypeInformation -Append
       }
      }
     }
     Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName" -Completed
     CheckOutput
    }
  7 {
      $Result=""  
      $Results=@() 
      $Path="$Location\Export all teams with Private Channels Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Write-Host Exporting teams with private Channels...
      $Count=0
      Get-Team | foreach {
       $TeamName=$_.DisplayName
       Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing Team: $TeamName "
       $Count++
       $GroupId=$_.GroupId
       $PrivateChannels=(Get-TeamChannel -GroupId $GroupId -MembershipType Private).DisplayName
       $PrivateChannelsCount=$PrivateChannels.count
       $PrivateChannelsName=$PrivateChannels -join ","
       if($PrivateChannelsCount -gt 0)
       {
        $Result=@{'Teams Name'=$TeamName;'Private Channels Count'=$PrivateChannelsCount; 'Private Channel Names'=$PrivateChannelsName}
        $Results= New-Object psobject -Property $Result
        $Results | select 'Teams Name','Private Channels Count','Private Channel Names' | Export-Csv $Path -NoTypeInformation -Append
       }
      }
      Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName  `n Currently Processing Channel: $ChannelName"  -Completed
      CheckOutput
     }

   }
   if($Action -ne "")
   {exit}
}
  While ($i -ne 0)
  Clear-Host
 