<#
=============================================================================================
Name:           Office 365 license reporting and management tool
Description:    This script can perform 10+ Office 365 reporting and management activities
Website:        o365reports.com

Script Highlights:
~~~~~~~~~~~~~~~~~ 

    1. Generates 5 Office 365 license reports. 
    2. Allows to perform 6 license management actions that includes, adding or removing licenses in bulk. 
    3. License Name is shown with its friendly name  like ‘Office 365 Enterprise E3’ rather than ‘ENTERPRISEPACK’.
    4. The script can be executed with an MFA enabled account too. 
    5. Exports the report result to CSV. 
    6. Exports license assignment and removal log file. 
    7. The script is scheduler-friendly. i.e., you can pass the credentials as a parameter instead of saving them inside the script.

For detailed Script execution: https://o365reports.com/2021/11/23/office-365-license-reporting-and-management-using-powershell
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$LicenseName,
    [string]$LicenseUsageLocation,
    [int]$Action,
    [switch]$MultipleActionsMode,
    [string]$UserName,
    [string]$Password
)

Function Open_OutputFile
{
 #Open output file after execution 
 if((Test-Path -Path $OutputCSVName) -eq "True")
 {
  if($ActionFlag -eq "Report")
  {
   Write-Host " Detailed license report is available in:" -NoNewline -ForegroundColor Yellow; Write-Host $OutputCSVName `n
   Write-Host The report has $ProcessedCount records
   Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
   Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
   Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
   }
  elseif($ActionFlag -eq "Mgmt")
  {
     Write-Host " License assignment/removal log file is available in:" -NoNewline -ForegroundColor Yellow; Write-Host $OutputCSVName 
     Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
     Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
     Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  } 
  $Prompt = New-Object -ComObject wscript.shell  
  $UserInput = $Prompt.popup("Do you want to open output file?",`  
 0,"Open Output File",4)  
  If ($UserInput -eq 6)  
  {  
   Invoke-Item "$OutputCSVName"  
  } 
 }
 Else
 {
  Write-Host `nNo records found
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
  Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 }
 Write-Progress -Activity Export CSV -Completed
}

#Get users' details
Function Get_UserInfo
{
 $global:DisplayName=$_.DisplayName
 $global:UPN=$_.UserPrincipalName
 $global:Licenses=$_.Licenses.AccountSkuId 
 $SigninStatus=$_.BlockCredential
 if($SigninStatus -eq $False)
 {$global:SigninStatus="Enabled"}
 else{$global:SigninStatus="Disabled"}
 $global:Department=$_.Department
 $global:JobTitle=$_.Title
 if($Department -eq $null)
 {$global:Department="-"}
 if($JobTitle -eq $null)
 {$global:JobTitle="-"}
}

Function Get_License_FriendlyName
{
 $FriendlyName=@()
 $LicensePlan=@()    
 #Convert license plan to friendly name 
 foreach($License in $Licenses) 
 {   
  $LicenseItem= $License -Split ":" | Select-Object -Last 1  
  $EasyName=$FriendlyNameHash[$LicenseItem]  
  if(!($EasyName))  
  {$NamePrint=$LicenseItem}  
  else  
  {$NamePrint=$EasyName} 
  $FriendlyName=$FriendlyName+$NamePrint
  $LicensePlan=$LicensePlan+$LicenseItem
 }
 $global:LicensePlans=$LicensePlan -join ","
 $global:FriendlyNames=$FriendlyName -join ","
}

Function Set_UsageLocation
{
 if($LicenseUsageLocation -ne "")
 {
  "Assigning Usage Location $LicenseUsageLocation to $UserName" |  Out-File $OutputCSVName -Append
  Set-MsolUser -UserPrincipalName $UserName -UsageLocation $LicenseUsageLocation
 }
 else
 {
  "Usage location is mandatory to assign license. Please set Usage location for $UserName" |  Out-File $OutputCSVName -Append
 }
}

Function Assign_Licenses
{
 "Assigning $LicenseNames license to $UPN" | Out-File $OutputCSVName -Append
 Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $LicenseNames
 if($?)
 {
  "License assigned successfully" | Out-File $OutputCSVName -Append
 }
 else
 {
  "License assignment failed" | Out-file $OutputCSVName -Append
 }
}

Function Remove_Licenses
{
 Write-Progress -Activity "`n     Removing $License license from $UPN "`n"  Processed users: $ProcessedCount"
 "Removing $License license from $UPN" | Out-File $OutputCSVName -Append
 Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $License
 if($?)
 {
  "License removed successfully" | Out-File $OutputCSVName -Append
 }
 else
 {
  "License removal failed" | Out-file $OutputCSVName -Append
 }
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
 $FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData
 [boolean]$Delay=$false


 Do {                 
  if($Action -eq "")
  {                       
                       
                
  Write-Host ""
  Write-host `nOffice 365 License Reporting -ForegroundColor Yellow
  Write-Host  "    1.Get all licensed users" -ForegroundColor Cyan
  Write-Host  "    2.Get all unlicensed users" -ForegroundColor Cyan
  Write-Host  "    3.Get users with specific license type" -ForegroundColor Cyan
  Write-Host  "    4.Get all disabled users with licenses" -ForegroundColor Cyan
  Write-Host  "    5.Office 365 license usage report" -ForegroundColor Cyan
  Write-Host `nOffice 365 License Management -ForegroundColor Yellow
  Write-Host  "    6.Bulk:Assign a license to users (input CSV)" -ForegroundColor Cyan
  Write-Host  "    7.Bulk:Assign multiple licenses to users (input CSV)" -ForegroundColor Cyan
  Write-Host  "    8.Remove all license from a user" -ForegroundColor Cyan
  Write-Host  "    9.Bulk:Remove all licenses from users (input CSV)" -ForegroundColor Cyan
  Write-Host  "    10.Remove specific license from all users" -ForegroundColor Cyan
  Write-Host  "    11.Remove all license from disabled users" -ForegroundColor Cyan
  Write-Host  "    0.Exit" -ForegroundColor Cyan
  Write-Host ""
  $GetAction = Read-Host 'Please choose the action to continue' 
 }
 else
 {
  $GetAction=$Action
 }

  Switch ($GetAction) {
   1 {
      $OutputCSVName=".\O365UserLicenseReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
      Write-Host `nGenerating licensed users report...`n
      $ProcessedCount=0
      Get-MsolUser -All | where {$_.IsLicensed -eq $true} | foreach {
       $ProcessedCount++
       Get_UserInfo
       Write-Progress -Activity "`n     Processed users count: $ProcessedCount "`n"  Currently Processing: $DisplayName"
       Get_License_FriendlyName
       $Result = @{'Display Name'=$Displayname;'UPN'=$upn;'License Plan'=$LicensePlans;'License Plan Friendly Name'=$FriendlyNames;'Account Status'=$SigninStatus;'Department'=$Department;'Job Title'=$JobTitle }
       $Results = New-Object PSObject -Property $Result
       $Results |select-object 'Display Name','UPN','License Plan','License Plan Friendly Name','Account Status','Department','Job Title' | Export-Csv -Path $OutputCSVName -Notype -Append
      }
      $ActionFlag="Report"
      Open_OutputFile
     }

  2 {
     $OutputCSVName=".\O365UnlicenedUserReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Write-Host `nGenerating Unlicensed users report...`n
     $ProcessedCount=0
     Get-MsolUser -All -UnlicensedUsersOnly | foreach {
      $ProcessedCount++
      Get_UserInfo
      Write-Progress -Activity "`n     Processed users count: $ProcessedCount "`n"  Currently Processing: $DisplayName"
      $Result = @{'Display Name'=$Displayname;'UPN'=$UPN;'Department'=$Department;'Signin Status'=$SigninStatus;'Job Title'=$JobTitle }
      $Results = New-Object PSObject -Property $Result
      $Results |select-object 'Display Name','UPN','Department','Job Title','Signin Status' | Export-Csv -Path $OutputCSVName -Notype -Append
     }
     $ActionFlag="Report"
     Open_OutputFile
    }

  3 {
     $OutputCSVName="./O365UsersWithSpecificLicenseReport__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     if($LicenseName -eq "")
     {
      $LicenseName=Read-Host "`nEnter the license SKU(Eg:contoso:Enterprisepack)"
     }
     Write-Host `nGetting users with $LicenseName license...`n
     $ProcessedCount=0
     if((Get-MsolAccountSku).AccountSkuID -icontains $LicenseName)
     {
      Get-MsolUser -All | Where-Object {($_.licenses).AccountSkuId -eq $LicenseName} | foreach{
       $ProcessedCount++
       Get_UserInfo
       Write-Progress -Activity "`n     Processed users count: $ProcessedCount "`n"  Currently Processing: $DisplayName"
       Get_License_FriendlyName
       $Result = @{'Display Name'=$Displayname;'UPN'=$upn;'License Plan'=$LicensePlans;'License Plan_Friendly Name'=$FriendlyNames;'Account Status'=$SigninStatus;'Department'=$Department;'Job Title'=$JobTitle }
       $Results = New-Object PSObject -Property $Result
       $Results |select-object 'Display Name','UPN','License Plan','License Plan_Friendly Name','Account Status','Department','Job Title' | Export-Csv -Path $OutputCSVName -Notype -Append
      }
     }
     else
     {
      Write-Host $LicenseName is not used in your organization. Please check the license name or run the License Usage Report to know the licenses in your org -ForegroundColor Red
     }
     #Clearing license name for next iteration
     $LicenseName=""
     $ActionFlag="Report"
     Open_OutputFile
    }

  4 {
     $OutputCSVName="./O365DiabledUsersWithLicense__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     $ProcessedCount=0
     Write-Host `nFinding disabled users still licensed in Office 365...`n
     Get-MsolUser -All -EnabledFilter DisabledOnly | where {$_.IsLicensed -eq $true} | foreach {
      $ProcessedCount++
      Get_UserInfo
      Write-Progress -Activity "`n     Processed users count: $ProcessedCount "`n"  Currently Processing: $DisplayName"
      $AssignedLicense="" 
      $FriendlyName=""
      $Count=0
      Get_License_FriendlyName
      $Result = @{'Display Name'=$Displayname;'UPN'=$upn;'License Plan'=$LicensePlans;'License Plan_Friendly Name'=$FriendlyNames;'Department'=$Department;'Job Title'=$JobTitle }
      $Results = New-Object PSObject -Property $Result
      $Results |select-object 'Display Name','UPN','License Plan','License Plan_Friendly Name','Department','Job Title' | Export-Csv -Path $OutputCSVName -Notype -Append
     }
     $ActionFlag="Report"
     Open_OutputFile
    }

  5 {
     $OutputCSVName="./Office365LicenseUsageReport__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Write-Host `nGenerating Office 365 license usage report...`n
     $ProcessedCount=0
     Get-MsolAccountSku | foreach {
      $ProcessedCount++
      $AccountSkuID=$_.AccountSkuID
      $LicensePlan= $_.SkuPartNumber
      $SubscriptionState=Get-MsolSubscription
      Write-Progress -Activity "`n     Retrieving license info "`n"  Currently Processing: $LicensePlan"
      $EasyName=$FriendlyNameHash[$LicensePlan]  
      if(!($EasyName))  
      {$FriendlyName=$LicenseItem}  
      else  
      {$FriendlyName=$EasyName} 
      $Result = @{'AccountSkuId'=$AccountSkuID;'License Plan_Friendly Name'=$FriendlyName;'Active Units'=$_.ActiveUnits;'Consumed Units'=$_.ConsumedUnits }
      $Results = New-Object PSObject -Property $Result
      $Results |select-object 'AccountSkuId','License Plan_Friendly Name','Active Units','Consumed Units' | Export-Csv -Path $OutputCSVName -Notype -Append
     }
     $ActionFlag="Report"
     Open_OutputFile
    }

  6 {
     $OutputCSVName="./Office365LicenseAssignment_Log__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
     $UserNamesFile=Read-Host "`nEnter the CSV file containing user names(Eg:D:/UserNames.csv)"
  
     #We have an input file, read it into memory
     $UserNames=@()
     $UserNames=Import-Csv -Header "UPN" $UserNamesFile
     $ProcessedCount=0
     $LicenseNames=Read-Host "`nEnter the license name(Eg:contoso:Enterprisepack)"
     Write-Host `nAssigning license to users...`n
     if((Get-MsolAccountSku).AccountSkuID -icontains $LicenseNames)
     {
      foreach($Item in $UserNames)
      {
       $ProcessedCount++
       $UPN=$Item.UPN
       Write-Progress -Activity "`n     Assigning $LicenseNames license to $UPN "`n"  Processed users: $ProcessedCount"
       $UsageLocation=(Get-MsolUser -UserPrincipalName $UPN).UsageLocation
       if($UsageLocation -eq $null)
       {
        Set_UsageLocation
       }
       else
       {
        Assign_Licenses
       }
      }
     }
     else
     {
      Write-Host $LicenseNames is not used in your organization. Please check the license name or run the License Usage Report to know the licenses in your org -ForegroundColor Red
     }
     #Clearing license name and input file location for next iteration
     $LicenseNames=""
     $UserNamesFile=""
     $ActionFlag="Mgmt"
     Open_OutputFile
    }

  7 {
     $OutputCSVName="./Office365LicenseAssignment_Log__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
     $UserNamesFile=Read-Host "`nEnter the CSV file containing user names(Eg:D:/UserNames.csv)"
    
     #We have an input file, read it into memory
     $UserNames=@()
     $UserNames=Import-Csv -Header "UPN" $UserNamesFile
     $Flag=""
     $ProcessedCount=0
     $LicenseNames=Read-Host "`nEnter the license names(Eg:TenantName:LicensePlan1,TenantName:LicensePlan2)"
     $LicenseNames=$LicenseNames.Replace(' ','')
     $LicenseNames=$LicenseNames.split(",")
     foreach($LicenseName in $LicenseNames)
     {  
      if((Get-MsolAccountSku).AccountSkuID -inotcontains $LicenseName)
      {
       $Flag="Terminate"
       Write-Host $LicenseName is not used in your organization. Please check the license name or run the License Usage Report to know the licenses in your org -ForegroundColor Red
      }
     }
     if($Flag -eq "Terminate")
     {
      Write-Host `nPlease re-run the script with appropriate license name -ForegroundColor Yellow
     }
     else
     {
      Write-Host `nAssigning licenses to Office 365 users...`n
      foreach($Item in $UserNames)
      {
       $UPN=$Item.UPN
       $ProcessedCount++
       $UsageLocation=(Get-MsolUser -UserPrincipalName $UPN).UsageLocation
       if($UsageLocation -eq $null)
       {
        Set_UsageLocation
       }
       else
       {
        Write-Progress -Activity "`n     Assigning licenses to $UPN "`n"  Processed users: $ProcessedCount"
        Assign_Licenses
       }
      }
     }
     #Clearing license names and input file location for next iteration
     $LicenseNames=""
     $UserNamesFile=""
     $ActionFlag="Mgmt"
     Open_OutputFile
    }
       

  8 {
     $Identity=Read-Host `nEnter User UPN
     $UserInfo=Get-MsolUser -UserPrincipalName $Identity
     #Checking whether the user is available
     if($UserInfo -eq $null)
     {
      Write-Host User $Identity does not exist. Please check the user name. -ForegroundColor Red
     }
     else
     {
      $Licenses=$UserInfo.Licenses.AccountSkuID
      Write-Host Removing $Licenses license from $Identity
      Set-MsolUserLicense -UserPrincipalName $Identity -RemoveLicenses $Licenses
      Write-Host `nAction completed -ForegroundColor Green
	  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
      Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
     }  
    }

  9 {
     $OutputCSVName="./Office365LicenseRemoval_Log__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
     $UserNamesFile=Read-Host "`nEnter the CSV file containing user names(Eg:D:/UserNames.csv)"
    
     #We have an input file, read it into memory
     $UserNames=@()
     $UserNames=Import-Csv -Header "UPN" $UserNamesFile
     $ProcessedCount=0
     foreach($Item in $UserNames)
     {
      $UPN=$Item.UPN
      $ProcessedCount++
     # Write-Progress -Activity "`n     Removing license from $UPN "`n"  Processed users: $ProcessedCount"
      $License=(Get-MsolUser -UserPrincipalName $UPN).licenses.AccountSkuID
      Remove_Licenses
     }
     $ActionFlag="Mgmt"
     Open_OutputFile 
    }

  10 {
      $OutputCSVName="./O365LicenseRemoval_Log__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
      $License=Read-Host "`nEnter the license name(Eg:TenantName:LicensePlan)"
      $ProcessedCount=0
      Write-Host `nRemoving $License license from users...`n
      if((Get-MsolAccountSku).AccountSkuID -icontains $License)
      {
       Get-MsolUser -All | Where-Object {($_.licenses).AccountSkuId -eq $License} | foreach{
        $ProcessedCount++
        $UPN=$_.UserPrincipalName
        Remove_Licenses
       }
      }
      else
      {
       Write-Host $License not used in your organization. Please check the license name or run the License Usage Report to know the licenses in your org -ForegroundColor Red
      }
      $ActionFlag="Mgmt"
      Open_OutputFile
     }  

  11 {
      $OutputCSVName="./O365LicenseRemoval_Log__$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
      Write-Host `nRemoving license from disabled users...`n
      $ProcessedCount=0
      Get-MsolUser -All -EnabledFilter DisabledOnly | where {$_.IsLicensed -eq $true} | foreach {
       $ProcessedCount++
       $UPN=$_.UserPrincipalName
       $License=$_.Licenses.AccountSkuID
       Remove_Licenses
      }
      $ActionFlag="Mgmt"
      Open_OutputFile
     } 
    }
    if($Action -ne "")
    {exit}
    if($MultipleActionsMode.ispresent)
    {                          
     Start-Sleep -Seconds 2
    } 
    else
    {
     Exit
    }
 }
  While ($GetAction -ne 0)
  Clear-Host
}
. main