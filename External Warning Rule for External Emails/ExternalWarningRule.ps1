#Create mail flow rule for adding External warning message for external mails
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$MFA,
    [string]$ExcludeGroupMembers,
    [string]$ExcludeMB,
    [string]$UserName,
    [string]$Password
)
Get-PSSession | Remove-PSSession
#Authentication using MFA
if($MFA.IsPresent)
{
 $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
 If ($MFAExchangeModule -eq $null)
 {
  Write-Host  `nPlease install Exchange Online MFA Module.  -ForegroundColor yellow
  Write-Host You can manually install module using below blog : `nhttps://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/ `nOR you can install module directly by entering "Y"`n
  $Confirm= Read-Host `nAre you sure you want to install module directly? [Y] Yes [N] No
  if($Confirm -match "[Y]")
  {
   Start-Process "iexplore.exe" "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"
  }
  else
  {
   Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
   Exit
  }
  $Confirmation= Read-Host Have you installed Exchange Online MFA Module? [Y] Yes [N] No
  if($Confirmation -match "[yY]")
  {
   $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
   If ($MFAExchangeModule -eq $null)
   {
    Write-Host Exchange Online MFA module is not available -ForegroundColor red
    Exit
   }
  }
  else
  {
   Write-Host Exchange Online PowerShell Module is required
   Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
   Exit
  }
 }

 #Importing Exchange MFA Module
 . "$MFAExchangeModule"
 Write-Host Enter credential in prompt to connect to Exchange Online
 Connect-EXOPSSession -WarningAction SilentlyContinue
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
 $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
 Import-PSSession $Session -AllowClobber -DisableNameChecking | Out-Null
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