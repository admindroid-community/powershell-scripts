<#
=============================================================================================
Name:           Configure External Email Warning Message for External Office 365 Emails
Version:        3.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~~

1.This script can be executed with MFA enabled account too. 
2.Prepends “External” to subject line for incoming external emails 
3.Adds “External Disclaimer” for external emails 
4.You can exclude group mailboxes like support, sales that facing external world. 

For detailed script execution:  https://o365reports.com/2020/03/25/how-to-add-external-email-warning-message/
============================================================================================
#>

#Create mail flow rule for adding External warning message for external mails
Param
(
    [Parameter(Mandatory = $false)]
    [string]$ExcludeGroupMembers,
    [string]$ExcludeMB,
    [string]$UserName,
    [string]$Password
)

#Check for EXO module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
   Import-Module ExchangeOnlineManagement
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
  Connect-ExchangeOnline -Credential $Credential
 }
 else
 {
  Connect-ExchangeOnline
 }

$Disclaimer='<p><div style="background-color:#FFEB9C; width:100%; border-style: solid; border-color:#9C6500; border-width:1pt; padding:2pt; font-size:10pt; line-height:12pt; font-family:Calibri; color:Black; text-align: left;"><span style="color:#9C6500"; font-weight:bold;>CAUTION:</span> This email originated from outside of the organization. Do not click links or open attachments unless you recognize the sender and know the content is safe.</div><br></p>'
$ExcludeGroups=@()
$ExcludeMBs=@()

#Check for exclude groups
if([string]$ExcludeGroupMembers -ne "")  
{  
 $Groups= $ExcludeGroupMembers -Split ","
 #Check whether the grup exists
 foreach($Group in $Groups)
 {
  $check=Get-DistributionGroup -Identity $Group -ErrorAction SilentlyContinue
  if($check -eq $null)
  {
   $check=Get-UnifiedGroup -Identity $Group -ErrorAction silentlycontinue
   if($check -eq $null)
   {
    Write-Host $Group not exist in the tenant -ForegroundColor Red
    continue
   }
  }
  $ExcludeGroups +=$Group
 }
 Write-Host `nCreating mail flow rule for External Senders... -ForegroundColor Yellow
 New-TransportRule "External Email Warning" -FromScope NotInOrganization -SentToScope InOrganization -PrependSubject [EXTERNAL]: -Priority 0 -ApplyHtmlDisclaimerText $Disclaimer -ExceptIfSentToMemberOf $ExcludeGroups -ApplyHtmlDisclaimerLocation Prepend -ApplyHtmlDisclaimerFallbackAction Wrap
}
#Create Transport rule with exclude mailbox that in To and CC
elseif($ExcludeMB -ne "")
{
 $MBs= $ExcludeMB -Split ","
 #Check whether the grup exists
 foreach($MB in $MBs)
 {
  $ExcludeMBs +=$MB
 }
 Write-Host `nCreating mail flow rule for External Senders... -ForegroundColor Yellow
 New-TransportRule "External Email Warning" -FromScope NotInOrganization -SentToScope InOrganization -PrependSubject [EXTERNAL]: -Priority 0 -ApplyHtmlDisclaimerText $Disclaimer -ExceptIfAnyOfToCCHeader $ExcludeMBs -ApplyHtmlDisclaimerLocation Prepend -ApplyHtmlDisclaimerFallbackAction Wrap
 }
else
{
Write-Host Creating mail flow rule for External Senders... -ForegroundColor Yellow
New-TransportRule "External Email Warning" -FromScope NotInOrganization -SentToScope InOrganization -PrependSubject [EXTERNAL]: -Priority 0 -ApplyHtmlDisclaimerText $Disclaimer -ApplyHtmlDisclaimerLocation Prepend -ApplyHtmlDisclaimerFallbackAction Wrap
}
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
Disconnect-ExchangeOnline