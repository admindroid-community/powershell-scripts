<#
=============================================================================================
Name:           Export calendar permission Report for Exchange Online Mailboxes
Version:        2.0
Website:        o365reports.com
    

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. Generates 6 different mailbox calendar permissions reports. 
2. The script uses modern authentication to retrieve calendar permissions. 
3. The script can be executed with MFA enabled account too.    
4. Exports report results to CSV file.    
5. Allows you to track all the calendars’ permission  
6. Helps to view default calendar permission for all the mailboxes 
7. Displays all the mailbox calendars to which a user has access. 
8. Lists calendars shared with external users. 
9. Helps to find out calendar permissions for a list of mailboxes through input CSV. 
10. Automatically install the EXO V2 module (if not installed already) upon your confirmation.   
11. The script is scheduler-friendly. I.e., Credential can be passed as a parameter instead of saving inside the script. 

For detailed Script execution: https://o365reports.com/2021/11/02/get-calendar-permissions-report-for-office365-mailboxes-powershell


Change Log
~~~~~~~~~~

    V1.0 (Nov 02, 2021) - File created
    V1.1 (Sep 28, 2023) - Minor changes
    V2.0 (Oct 14, 2024) - Updated the script to use REST based cmdlets and added certificate-based authentication support to enhance scheduling capability
    
============================================================================================
#>
param (
    [string] $UserName = $null,
    [string] $Password = $null,
    [Switch] $ShowAllPermissions,
    [String] $DisplayAllCalendarsSharedTo,
    [Switch] $DefaultCalendarPermissions,
    [Switch] $ExternalUsersCalendarPermissions,
    [String] $CSVIdentityFile,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint
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
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
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


Function OutputFile_Declaration
{
 if ($DisplayAllCalendarsSharedTo -ne "")
 {
  $global:ExportCSVFileName = "CalendarsSharedTo"+ $DisplayAllCalendarsSharedTo+"_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
 }
 elseif ($ShowAllPermissions.IsPresent) 
 {
  $global:ExportCSVFileName = "AllCalendarPermissionReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
 }
 elseif ($DefaultCalendarPermissions.IsPresent) 
 {
  $global:ExportCSVFileName = "DefaultCalendarPermissionsReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
 }
 elseif ($ExternalUsersCalendarPermissions.IsPresent) 
 {
  $global:ExportCSVFileName = "SharedCalendarsWithExternalUsersReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
 }
 else 
 {
  $global:ExportCSVFileName = "CalendarPermissionsReport-" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv" 
 }
}

#Checks the user input file availability. Then, processes the mailbox 
Function RetrieveMBs 
{
 if ([string]$CSVIdentityFile -ne "") 
 {
  $IdentityList = Import-Csv -Header "IdentityValue" $CSVIdentityFile
  foreach ($Identity in $IdentityList) {
   $CurrIdentity = $Identity.IdentityValue
   $CurrUserData = Get-EXOMailbox -identity $currIdentity -ErrorAction SilentlyContinue 
   if ($null -eq $CurrUserData) 
   {
    Write-Host $currIdentity mailbox is not found/invalid.
   }
   else 
   {
    GetCalendars                 
   }
  }
 }
 else 
 {
  Get-EXOMailbox -ResultSize Unlimited | ForEach-Object {
   $CurrUserData = $_
   GetCalendars
  }
 }
}

Function GetCalendars
{
 $global:MailboxCount = $global:MailboxCount+1
 $EmailAddress = $CurrUserData.PrimarySmtpAddress
 $global:DisplayName=$CurrUserData.DisplayName
 $CalendarFolders=@()
 $CalendarStats = Get-EXOMailboxFolderStatistics -Identity $EmailAddress -FolderScope Calendar
 
 #Processing the calandar folder path
 ForEach($LiveCalendarFolder in $CalendarStats) 
 {
  if (($LiveCalendarFolder.FolderType) -eq "Calendar") 
  {
   $CurrCalendarFolder = $EmailAddress + ":\Calendar"
  }
  else 
  {
   $CurrCalendarFolder = $EmailAddress + ":\Calendar\" + $LiveCalendarFolder.Name
  }
  $CalendarFolders += $CurrCalendarFolder
  
 }
 RetrieveCalendarPermissions
}

#Processes the mailbox calendars and retrieves the calendar permissions 
Function RetrieveCalendarPermissions 
{ 
 #Determine the usecase 
     
 #Processing the DisplayAllCalendarsSharedTo switch param   
 if ($DisplayAllCalendarsSharedTo -ne "") 
 {
  $DisplayName=$CurrUserData.DisplayName
  $Flag = "DisplayAllCalendarsSharedTo"
  foreach($CalendarFolder in $CalendarFolders)
  {
   $CalendarName=$CalendarFolder -split "\\" | Select-Object -Last 1
   Write-Progress "Checking calendar permission in: $CalendarFolder" "Processed mailbox count: $global:MailboxCount"
   $CurrCalendarData = Get-EXOMailboxFolderPermission -Identity $CalendarFolder -User $CurrMailboxData.PrimarySmtpAddress -ErrorAction SilentlyContinue 
   if ($null -ne $CurrCalendarData) 
   {
    SaveCalendarPermissionsData
   }
  }
 }

 #Processing the ShowAllPermissions switch param  
 elseif ($ShowAllPermissions.IsPresent) 
 {
  foreach($CalendarFolder in $CalendarFolders)
  {
   $CalendarName=$CalendarFolder -split "\\" | Select-Object -Last 1
   Write-Progress "Checking calendar permission in: $CalendarFolder" "Processed mailbox count: $global:MailboxCount"
   Get-EXOMailboxFolderPermission -Identity $CalendarFolder | foreach {
    $CurrCalendarData=$_
    SaveCalendarPermissionsData
   }
  }
 }

 #Processing the DefaultCalendarPermissions switch param 
 elseif ($DefaultCalendarPermissions.IsPresent) 
 {
  $Flag = "DefaultUserCalendar"
  foreach($CalendarFolder in $CalendarFolders)
  {
   Write-Progress "Checking default calendar permission for $CalendarFolder" "Processed mailbox count: $global:MailboxCount"
   $CalendarName=$CalendarFolder -split "\\" | Select-Object -Last 1
   $CurrCalendarData= Get-EXOMailboxFolderPermission -Identity $CalendarFolder | where-Object { $_.User.ToString() -eq "Default" } #| foreach-object {
   SaveCalendarPermissionsData
  } 
 }

 #Processing the ExternalUsersCalendarPermissions switch param 
 elseif ($ExternalUsersCalendarPermissions.IsPresent) 
 {
  $Flag = "ExternalUserCalendarSharing"
  foreach($CalendarFolder in $CalendarFolders)
  {
   Write-Progress "Checking default calendar permission for $CalendarFolder" "Processed mailbox count: $global:MailboxCount"
   $CalendarName=$CalendarFolder -split "\\" | Select-Object -Last 1
   Get-EXOMailboxFolderPermission -Identity $CalendarFolder | where-Object { $_.User.DisplayName.StartsWith("ExchangePublishedUser.") }  | foreach-object {
   $CurrCalendarData=$_
   SaveCalendarPermissionsData
   }
  }
 }

 #Processing default report
 else 
 {
  foreach($CalendarFolder in $CalendarFolders)
  {
   Write-Progress "Checking calendar permission for $CalendarFolder" "Processed mailbox count: $global:MailboxCount"
   $CalendarName=$CalendarFolder -split "\\" | Select-Object -Last 1
   Get-EXOMailboxFolderPermission -Identity $CalendarFolder | where-Object { ($_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous")  }  | foreach-object {
    $CurrCalendarData=$_
	SaveCalendarPermissionsData
   }
  }
 }
}

Function SaveCalendarPermissionsData 
{
 $Identity = $CurrUserData.Identity
 $global:ReportSize = $global:ReportSize + 1
 $MailboxType = $CurrUserData.RecipientTypeDetails
 $CalendarName = $CalendarName
 $SharedToMB=$CurrCalendarData.User.DisplayName
 if ($SharedToMB.StartsWith("ExchangePublishedUser.")) 
 {
  $AllowedUser = $SharedToMB -replace ("ExchangePublishedUser.", "")      
  $UserType = "External/Unauthorized"
 }
 else 
 {
  $AllowedUser = $SharedToMB
  $UserType = "Member"
 }
 $AccessRights = $CurrCalendarData.AccessRights -join ","
 if ($Empty -ne ($CurrCalendarData.SharingPermissionFlags)) 
 {
  $PermissionFlag = $CurrCalendarData.SharingPermissionFlags -join ","
 }
 else 
 {
  $PermissionFlag = "-"
 }
 ExportCalendarPermissionData
}

#Exporting all the processed data to the appropriate switch params of user choice
Function ExportCalendarPermissionData 
{ 
    $ExportResult = @{ 
        'Mailbox Name'             = $Identity;
        'Email Address'            = $EmailAddress;
        'Mailbox Type'             = $MailboxType; 
        'Calendar Name'            = $CalendarName;
        'Shared To'                = $AllowedUser;
        'User Type'                = $UserType;
        'Access Rights'            = $AccessRights;
        'Sharing Permission Flags' = $PermissionFlag;
    }
   
    
 $ExportResults = New-Object PSObject -Property $ExportResult
 if ($Flag -eq "DisplayAllCalendarsSharedTo") 
 {
  $ExportResults | Select-object 'Mailbox Name', 'Email Address','Calendar Name', 'Access Rights','Sharing Permission Flags', 'Mailbox Type' | Export-csv -path $global:ExportCSVFileName -NoType -Append
 }

 elseif ($Flag -eq "DefaultUserCalendar") 
 {
  $ExportResults | Select-object 'Mailbox Name', 'Email Address', 'Mailbox Type', 'Calendar Name', 'Access Rights' | Export-csv -path $global:ExportCSVFileName -NoType -Append
 }
 elseif ($Flag -eq "ExternalUserCalendarSharing")
 {
  $ExportResults | Select-object 'Mailbox Name',  'Email Address', 'Calendar Name', 'Shared To' ,'Access Rights'  | Export-csv -path $global:ExportCSVFileName -NoType -Append
 }
 else 
 {
  $ExportResults | Select-object 'Mailbox Name', 'Email Address', 'Mailbox Type', 'Calendar Name', 'Shared To', 'Access Rights', 'Sharing Permission Flags', 'User Type' | Export-csv -path $ExportCSVFileName -NoType -Append
 }
}

#Execution starts here
Connect_Exo
OutputFile_Declaration
$global:MailboxCount = 0
$global:ReportSize = 0
if ($DisplayAllCalendarsSharedTo -ne "") 
{
 $CurrMailboxData = Get-EXOMailbox -Identity $DisplayAllCalendarsSharedTo -ErrorAction SilentlyContinue 
 if ($CurrMailboxData -eq $null) 
 {
  Write-Host "Given email address is invalid. Exiting from execution." -ForegroundColor Magenta
  return
 }
}  
Write-Host Generating mailboxes"'" calendar permission report...
RetrieveMBs

#Validates the output file availability
if ((Test-Path -Path $ExportCSVFileName) -eq "True") { 
    #Open file after code execution finishes   
    Write-Host `n" The output file available in:" -NoNewline -ForegroundColor Yellow; Write-Host .\$global:ExportCSVFileName `n
    write-host "Exported $global:ReportSize records to CSV." `n
    Write-Host "Disconnected active ExchangeOnline session" `n 
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline;
    Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n     
    $prompt = New-Object -ComObject wscript.shell    
    $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
    If ($userInput -eq 6) {    
        Invoke-Item "$global:ExportCSVFileName"
    }  
} 
else {
    Write-Host "No data found with the specified criteria"
}

#Disconneting the ExchangeOnline connection
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue  