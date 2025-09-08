<#
=============================================================================================
Name:           Export Exchange Online room mailbox reports
Description:    A single script can generate 8 detailed room mailbox reports
Version:        1.0
Website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1.A single script can generate 8 Room mailbox reports.    
2.The script can be executed with an MFA enabled account too.
3.The script supports certificate based authentication (CBA)       
4.Exports report results to CSV file.    
5.Lists all mailboxes and their capacity
6.Helps to export meeting room booking details
7.Helps to identify room mailboxes' resource delegates
8.Exports room mailbox permission details, including full access, sendas, and sendonbehalf permission 
9.With built-in filtering options, more granular reports can be generated
    -Meeting rooms that any one can book
    -Meeting rooms that can allow only specific persons to book meetings
    -List meeting rooms that require approval
    -Meeting rooms that can book by external users    
10.Automatically installs the EXO module (if not installed already) upon your confirmation.   
11.The script is scheduler friendly.

For detailed script execution: https://o365reports.com/2023/12/28/export-microsoft-365-room-mailbox-reports-using-powershell/
============================================================================================
#>



param(
[string]$Organization,
[string]$ClientId,
[string]$CertificateThumbprint,
[string]$CSVFilePath,
[Switch]$AnyoneCanBook,
[Switch]$BookingAllowedForLimitedPersons,
[Switch]$RequiresApproval,
[Switch]$AllowsBookingForExternalUsers,
[string]$UserName,
[string]$Password,
[int]$Action
)

Function Connect_Exo
{
 #Check for EXO module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell  module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
  } 
  else 
  { 
   Write-Host EXO module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 
 Write-Host Connecting to Exchange Online...
 #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
 }
 elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
   Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
 }
 else
 {
  Connect-ExchangeOnline -ShowBanner:$false
 }
}
Connect_Exo

if($Action -eq "")
{ 
 Write-Host ""
 Write-Host  "    1.Get all room mailboxes and their capacity" -ForegroundColor Cyan
 Write-Host  "    2.Export room mailboxes' booking options" -ForegroundColor Cyan
 Write-Host  "    3.Get room mailbox booking delegates" -ForegroundColor Cyan
 Write-Host  "    4.Get room mailbox permissions" -ForegroundColor Cyan
 Write-Host ""
 $GetAction = Read-Host 'Please choose the action to continue' 
 }
 else
 {
  $GetAction=$Action
 }
  $Result = ""  
    $Results = @() 
  $Count=0
 Switch ($GetAction) {
  1 {
     $Path="./RoomMailboxReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-Mailbox -ResultSize unlimited -RecipientTypeDetails "RoomMailbox" | foreach {
      $Count++
      $UPN=$_.UserPrincipalName
      $Name=$_.DisplayName
      $PrimarySMTPAddress=$_.PrimarySMTPAddress
      $Alias=$_.alias
      $Capacity=$_.ResourceCapacity
      Write-Progress -Activity "`n     Processing room: $Count - $UPN"
      $Result=@{'Room Mailbox Name'=$Name;'UPN'=$UPN;'Primary SMTP Address'=$PrimarySMTPAddress;'Alias'=$Alias;'Capacity'=$Capacity}
      $Results= New-Object psobject -Property $Result
      $Results | select 'Room Mailbox Name','UPN','Primary SMTP Address','Alias','Capacity' | Export-Csv $Path -NoTypeInformation -Append
     }
     Write-Host `nThe output file contains $Count room mailbox records
    }

  2 {
     $Path="./RoomMailbox_BookingOptionsReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     $OutputCount=0
     if(($AnyoneCanBook.IsPresent) -or ($BookingAllowedForLimitedPersons.IsPresent) -or ($RequiresApproval.IsPresent) -or ($AllowsBookingForExternalUsers.IsPresent))
     {
      $FilterPresent = 'True'
     }
     else
     {
      $FilterPresent = 'False'
     }
     Get-Mailbox -ResultSize unlimited -RecipientTypeDetails "RoomMailbox" | foreach {
      $Count++
      $Print=0
      $UPN=$_.UserPrincipalName
      $Name=$_.DisplayName
      $Capacity=$_.ResourceCapacity
      $ResourceDelegates=$RoomDetails.ResourceDelegates
      $Delegates=$ResourceDelegates -join ","
      Write-Progress -Activity "`n     Processing room: $Count - $UPN"
      $BookingDetails=Get-CalendarProcessing -Identity $UPN
      $ResourceDelegates=$BookingDetails.ResourceDelegates
      $Delegates=$ResourceDelegates -join ","
      $AllBookInPolicy=$BookingDetails.AllBookInPolicy
      $AllRequestInPolicy=$BookingDetails.AllRequestInPolicy
      $AllRequestOutOfPolicy=$BookingDetails.AllRequestOutOfPolicy
      $BookInPolicy=$BookingDetails.BookInPolicy
      $Process_BookInPolicy=""
      if($BookInPolicy -ne "")
      { 
       foreach ($obj in $BookInPolicy){
        $output = ($obj -split '-' | Select-Object -Skip 1) -join '-'
        if($Process_BookInPolicy -ne "")
        { 
         $Process_BookInPolicy=$Process_BookInPolicy + ','
        }
        $Process_BookInPolicy += $output 
       }
      }
      $BookInPolicy=$Process_BookInPolicy

      $RequestInPolicy=$BookingDetails.RequestInPolicy
      $Process_RequestInPolicy=""
      if($RequestInPolicy -ne "")
      { 
       foreach ($obj in $RequestInPolicy){   
        $output = ($obj -split '-' | Select-Object -Skip 1) -join '-'
        if($Process_RequestInPolicy -ne "")
        { 
         $Process_RequestInPolicy=$Process_RequestInPolicy + ','
        }
        $Process_RequestInPolicy += $output 
       }
      }
      $RequestInPolicy=$Process_RequestInPolicy
      

      $RequestOutOfPolicy=$BookingDetails.RequestOutOfPolicy
      #$RequestOutOfPolicy = ($RequestOutOfPolicy | ForEach-Object { $_ -split '-' | Select-Object -Last 1 }) -join ','
      $Process_RequestOutOfPolicy=""
      if($RequestOutOfPolicy -ne "")
      { 
       foreach ($obj in $RequestOutOfPolicy){
        $output = ($obj -split '-' | Select-Object -Skip 1) -join '-'
        if($Process_RequestOutOfPolicy -ne "")
        { 
         $Process_RequestOutOfPolicy=$Process_RequestOutOfPolicy + ','
        }
        $Process_RequestOutOfPolicy += $output 
       }
      }
      $RequestOutOfPolicy=$Process_RequestOutOfPolicy

      $BookingWindow=$BookingDetails.BookingWindowInDays
      $MaximumDuration=$BookingDetails.MaximumDurationInMinutes
      $MinimumDuration=$BookingDetails.MinimumDurationInMinutes
      $AllowConflicts=$BookingDetails.AllowConflicts
      $AllowRecurringMeetings=$BookingDetails.AllowRecurringMeetings
      $EnforceCapacity=$BookingDetails.EnforceCapacity
      $AutomateProcessing=$BookingDetails.AutomateProcessing
      $ProcessExternalMeetingMessages=$BookingDetails.ProcessExternalMeetingMessages
      if($FilterPresent -eq 'False')
      {
       $Print=1
      }
      else
      {
       if($AnyoneCanBook.IsPresent -and ($AllBookInPolicy -eq $true))
       {
        $Print=1
       }
       elseif($BookingAllowedForLimitedPersons.IsPresent -and (($BookInPolicy -ne "") -and ($AllBookInPolicy -eq $false)))
       {
        $Print=1
       }
       elseif($RequiresApproval.IsPresent -and (($AllBookInPolicy -eq $false) -and ($AllRequestInPolicy -eq $true) -and ($ResourceDelegates -ne "")))
       {
        $Print=1
       }
       elseif($AllowsBookingForExternalUsers.IsPresent -and $ProcessExternalMeetingMessages -eq $true)
       {
        $Print=1
       }
      }
      if($BookInPolicy -eq "")
      { $BookInPolicy="-" }
      if($RequestinPolicy -eq "")
      { $RequestInPolicy="-" }
      if($RequestOutOfPolicy -eq "")
      { $RequestOutOfPolicy="-" }
      if($Delegates -eq "")
      { $Delegates="-" }
      if($Print -eq 1)
      {
       $OutputCount++
       $Result=@{'Room Mailbox Name'=$Name;'UPN'=$UPN;'Capacity'=$Capacity;'All Book In Policy'=$AllBookInPolicy;'All Request In Policy'=$AllRequestInPolicy;'All Request Out Of Policy'=$AllRequestOutOfPolicy;'Resource Delegate'=$Delegates;'Book In Policy'=$BookInPolicy;'Request In Policy'=$RequestInPolicy;'Request Out Of Policy'=$RequestOutOfPolicy;'Booking Window (days)'=$BookingWindow;'Max Duration (mins)'=$MaximumDuration;'Min Duration (mins)'=$MinimumDuration;'Allow Booking for External Users'=$ProcessExternalMeetingMessages;'Allow Conflicts'=$AllowConflicts;'Allow Recurring Meetings'=$AllowRecurringMeetings;'Enforce Capacity'=$EnforceCapacity}
       $Results= New-Object psobject -Property $Result
       $Results | select 'Room Mailbox Name','UPN','Capacity','All Book In Policy','All Request In Policy','All Request Out Of Policy','Resource Delegate','Book In Policy','Request In Policy','Request Out Of Policy','Booking Window (days)','Max Duration (mins)','Min Duration (mins)','Allow Booking for External Users','Allow Conflicts','Allow Recurring Meetings','Enforce Capacity' | Export-Csv $Path -NoTypeInformation -Append
      }
     }
     Write-Host `nThe output file contains $OutputCount room mailbox records
    }

    3 {
     $Path="./RoomMailboxDelegates_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-Mailbox -ResultSize unlimited -RecipientTypeDetails "RoomMailbox" | foreach {
      $Count++
      $UPN=$_.UserPrincipalName
      $Name=$_.DisplayName
      $PrimarySMTPAddress=$_.PrimarySMTPAddress
      $RoomDetails=Get-CalendarProcessing -Identity $UPN
      $ResourceDelegates=$RoomDetails.ResourceDelegates
      $Delegates=$ResourceDelegates -join ","
      if($Delegates -eq "")
      {
       $Delegates="-"
      }
      Write-Progress -Activity "`n     Processing room: $Count - $UPN"
      $Result=@{'Room Mailbox Name'=$Name;'UPN'=$UPN;'Primary SMTP Address'=$PrimarySMTPAddress;'Resource Delegates'=$Delegates}
      $Results= New-Object psobject -Property $Result
      $Results | select 'Room Mailbox Name','UPN','Primary SMTP Address','Resource Delegates' | Export-Csv $Path -NoTypeInformation -Append
     }
     Write-Host `nThe output file contains $Count room mailbox records
    }

  4 {
     $Path="./RoomMailbox_BookingOptionsReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
     Get-Mailbox -ResultSize unlimited -RecipientTypeDetails "RoomMailbox" | foreach {
      $Count++
      $UPN=$_.UserPrincipalName
      $Name=$_.DisplayName
      $SendOnBehalf=$_.GrantSendOnBehalfTo
      $SendOnBehalf=$SendOnBehalf -join ","
      $SendAs=(Get-RecipientPermission -Identity $UPN | Where{ -not (($_.Trustee -match "NT AUTHORITY") -or ($_.Trustee -match "S-1-5-21"))}).Trustee
      $SendAs=$SendAs -join ","
      $FullAccess=(Get-EXOMailboxPermission -Identity $UPN | Where { ($_.AccessRights -contains "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -match "NT AUTHORITY" -or $_.User -match "S-1-5-21") }).User
      $FullAccess=$FullAccess -join ","
      if($SendOnBehalf -eq "")
      { $SendOnBehalf="-"}
      if($SendAs -eq "")
      { $SendAs="-"}
      if ($FullAccess -eq "")
      { $FullAccess="-"}
      Write-Progress -Activity "`n     Processing room: $Count - $RoomAddress"
      $Result=@{'Room Mailbox Name'=$Name;'UPN'=$UPN;'Full Access'=$FullAccess;'Send As'=$SendAs;'Send On Behalf'=$SendOnBehalf}
      $Results= New-Object psobject -Property $Result
      $Results | select 'Room Mailbox Name','UPN','Full Access','Send As','Send On Behalf' | Export-Csv $Path -NoTypeInformation -Append
     }
     Write-Host `nThe output file contains $Count room mailbox records
    }
   }


  #Open output file after execution
 if((Test-Path -Path $Path) -eq "True") 
 {
  Write-Host `n" The Output file available in "  -NoNewline -ForegroundColor Yellow; Write-Host $Path 
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green  
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; 
  Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$Path"   
  } 
 }


#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false

