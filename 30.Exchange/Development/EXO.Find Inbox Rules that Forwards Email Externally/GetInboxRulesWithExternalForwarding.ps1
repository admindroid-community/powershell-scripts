<#
=============================================================================================
Name:           Find inbox rules with external email forwarding
Description:    This script Finds all inbox rules that forwards emails externally in Office 365 using PowerShell 
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.The script uses modern authentication to connect to Exchange Online.
2.The script can be executed with MFA enabled account.
3.Automatically installs the EXO V2 module (if not installed already) upon your confirmation.
4.Helps to filter-out forwarding rules that forwards emails to external users by excluding guest accounts. 
5.The script is scheduler – friendly. i.e., credentials can be passed as parameters rather than being saved inside the script.
6.Exports the report result to a CSV file. 

For detailed script execution: https://o365reports.com/2022/06/09/find-office365-inbox-rules-with-external-forwarding-powershell
============================================================================================
#>

#PARAMETERS
param ( 
[string] $UserName = $null, 
[String] $Password = $null,
[Switch] $ExcludeGuestUsers
) 

#Check for ExchangeOnline module availability
$Exchange = (Get-Module ExchangeOnlineManagement -ListAvailable).Name
if ($Exchange -eq $null)
{
 Write-Host "Important: ExchangeOnline PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."  
 $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
 if ($Confirm -match "[yY]") 
 { 
  Write-Host "Installing ExchangeOnlineManagement..." -ForegroundColor Magenta
  Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
  Import-Module ExchangeOnlineManagement -Force
  Write-Host "ExchangeOnline PowerShell module is installed in the machine successfully." -ForegroundColor Green `n
 }
 else
 { 
  Write-Host "Exiting. `nNote: ExchangeOnline PowerShell module must be available in your system to run the script." 
  Exit 
 }
}

#Connecting to ExchangeOnline .......
Write-Host "Connecting to ExchangeOnline..." -ForegroundColor Cyan
if(($UserName -ne "") -and ($Password -ne ""))
{
 $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
 $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
 Connect-ExchangeOnline -Credential $Credential
}
else
{
 Connect-ExchangeOnline
}
Write-Host "ExchangeOnline module successfully connected..." -ForegroundColor Green `n

#Function for export output to CSV file
Function ExportCSV
{
 $ForwardTo = if($ForwardTo.External.count -eq 0 -and $ForwardTo.Guest.count -eq 0){"-"}else{($ForwardTo.External + $ForwardTo.Guest)-join ', '}
 $ForwardAsAttachmentTo = if($ForwardAsAttachmentTo.External.count -eq 0 -and $ForwardAsAttachmentTo.Guest.count -eq 0){"-"}else{($ForwardAsAttachmentTo.External + $ForwardAsAttachmentTo.Guest) -join ', '}
 $RedirectTo = if($RedirectTo.External.count -eq 0 -and $RedirectTo.Guest.count -eq 0){"-"}else{($RedirectTo.External + $RedirectTo.Guest) -join ', '}
 if($ForwardTo -ne "-" -or $ForwardAsAttachmentTo -ne "-" -or $RedirectTo -ne "-")
 {
  $Global:InboxRuleWithExternalEmailCount++
  $Result = @{'Mailbox Name'= $MailBoxName;'UPN'= $UPN;'Inbox Rule Name'= $_.Name;'Forward To'= $ForwardTo; 'Forward As Attachment To'= $ForwardAsAttachmentTo; 'Redirect To'= $RedirectTo}  
  $ExportResult = New-Object PSObject -Property $Result
  $ExportResult | Select-Object 'Mailbox Name','UPN','Inbox Rule Name','Forward To','Forward As Attachment To','Redirect To' | Export-CSV $OutputCsv  -NoTypeInformation -Append 
 }
}

#Getting InboxRules with external email configuration......
Write-Host "Getting InboxRules with external email configuration..."`n
$OutputCsv=".\InboxRuleWithExternalEmails_$((Get-Date -format MMM-dd` hh-mm` tt).ToString()).csv"
$InboxRuleCount = 0
$Global:InboxRuleWithExternalEmailCount = 0
$GuestUser = Get-MailUser -ResultSize unlimited

Get-Mailbox -ResultSize unlimited | foreach {
 
 $MailBoxName = $_.DisplayName
 $UPN = $_.UserPrincipalName
 Write-Progress -Activity "Updating...    $Global:InboxRuleWithExternalEmailCount InboxRule found with external email " -Status "Getting InboxRule from '' $MailBoxName'' mailbox"
 Get-InboxRule -MailBox $_.PrimarySmtpAddress | foreach {
  
  $InboxRuleCount++
  $ForwardTo = @{External=@();Guest=@()}
  $ForwardAsAttachmentTo = @{External=@();Guest=@()}
  $RedirectTo = @{External=@();Guest=@()}

  #ForwardTo
  if ($_.ForwardTo -ne $null )
  {
   $UserinForwardTo = ( $_.ForwardTo -split ',' )
   $ExternalUser= $UserinForwardTo | Where-Object {$_ -match "SMTP:"}
   $InternalAndGuestUser = $UserinForwardTo | Where-Object {$_ -notmatch "SMTP:"}

   #Get external users email in ForwardTo
   foreach($ExternalUsers in $ExternalUser)
   {
    $ForwardTo.External += (($ExternalUsers -split " ")[0].TrimStart('"')).TrimEnd('"') 
   }

   #Get guest users email in ForwardTo
   if(!$ExcludeGuestUsers.IsPresent)
   {
    foreach($InternalAndGuestUsers in $InternalAndGuestUser)
    {
     $InternalAndGuestUsers= (($InternalAndGuestUsers -split ":")[1])
     $GuestUser | foreach {
     
      if($InternalAndGuestUsers -eq $_.LegacyExchangeDN+"]")
      {
       $ForwardTo.Guest += $_.OtherMail
      }
     }
    }
   }
  }

  #ForwardAsAttachmentTo
  if ($_.ForwardAsAttachmentTo -ne $null )
  {
   $UserinForwardAsAttachmentTo = ( $_.ForwardAsAttachmentTo -split ',' )
   $ExternalUser = $UserinForwardAsAttachmentTo | Where-Object {$_ -match "SMTP:"}
   $InternalAndGuestUser = $UserinForwardAsAttachmentTo | Where-Object {$_ -notmatch "SMTP:"}

   #Get external users email in ForwardAsAttachmentTo
   foreach($ExternalUsers in $ExternalUser)
   {
    $ForwardAsAttachmentTo.External += (($ExternalUsers -split " ")[0].TrimStart('"')).TrimEnd('"') 
   }

   #Get guest users email in ForwardAsAttachmentTo
   if(!$ExcludeGuestUsers.IsPresent)
   {
    foreach($InternalAndGuestUsers in $InternalAndGuestUser)
    {
     $InternalAndGuestUsers= (($InternalAndGuestUsers -split ":")[1])
     $GuestUser | foreach {
     
      if($InternalAndGuestUsers -eq $_.LegacyExchangeDN+"]")
      {
       $ForwardAsAttachmentTo.Guest += $_.OtherMail
      }
     }
    }
   }
  }

  #RedirectTo
  if ($_.RedirectTo -ne $null )
  {
   $UserinRedirectTo = ( $_.RedirectTo -split ',' )
   $ExternalUser = $UserinRedirectTo | Where-Object {$_ -match "SMTP:"}
   $InternalAndGuestUser = $UserinRedirectTo | Where-Object {$_ -notmatch "SMTP:"}

   #Get external users email in RedirectTo
   foreach($ExternalUsers in $ExternalUser)
   {
    $RedirectTo.External += (($ExternalUsers -split " ")[0].TrimStart('"')).TrimEnd('"') 
   }

   #Get guest users email in RedirectTo
   if(!$ExcludeGuestUsers.IsPresent)
   {
    foreach($InternalAndGuestUsers in $InternalAndGuestUser)
    {
     $InternalAndGuestUsers= (($InternalAndGuestUsers -split ":")[1])
     $GuestUser | foreach {
     
      if($InternalAndGuestUsers -eq $_.LegacyExchangeDN+"]")
      {
       $RedirectTo.Guest += $_.OtherMail
      }
     }
    }
   }
  }

  #Export output to CSV file
  ExportCSV
 }
}
Write-Host "Script executed successfully." `n

#Open output file after execution
if($InboxRuleCount -eq 0)
{
 Write-Host "No InboxRule found in this organization."
}
else
{
 Write-Host "$InboxRuleCount InboxRule found in this organization." `n
 if($Global:InboxRuleWithExternalEmailCount -eq 0)
 {
  Write-Host "But no InboxRule found with external email."
 }
 else
 {
  Write-Host "$Global:InboxRuleWithExternalEmailCount InboxRule found with external email."
  if((Test-Path -Path $OutputCsv) -eq "True") 
  {   
   Write-Host `n "The InboxRule with external email configuration report available in:" -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCsv" 
   $Prompt = New-Object -ComObject wscript.shell    
   $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)    
   If ($UserInput -eq 6)    
   {    
    Invoke-Item "$OutputCSV"    
   } 
 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n   
  }
 }
}
  


#Clean up session
Get-PSSession | Remove-PSSession