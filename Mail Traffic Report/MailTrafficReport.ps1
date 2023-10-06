<#
=============================================================================================
Name:           Export Office 365 mail traffic statistics by user report
Description:    This script exports mails sent, received, spam received and malware received statistics by users to CSV file
Version:        3.0
Website:        o365reports.com

Script Highlights:   
~~~~~~~~~~~~~~~~~
1.The script can generate 5+ email statistics reports like emails sent, emails received, spam received, and malware received count.  
2.The script uses modern authentication to connect to Exchange Online.   
3.The script can be executed with MFA enabled account too.   
4.Exports report results to CSV.   
5.Allows you to generate email statistics reports for a custom period.   
6.Automatically installs the EXO V2 module (if not installed already) upon your confirmation.   
7.Allows you to filter the mail traffic report for organization users alone.   
8.The script is scheduler-friendly. i.e., Credentials can be passed as a parameter.  

For detailed script execution: https://o365reports.com/2020/08/12/export-office-365-mail-traffic-report-with-powershell/
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [switch]$NoMFA,
    [switch]$OnlyOrganizationUsers,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [switch]$MailsSent,
    [switch]$SpamsReceived,
    [switch]$MalwaresReceived,
    [string]$UserName,
    [string]$Password
)

Function Install_Modules
{
 #Check for EXO v2 module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell V2 module is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
   Import-Module ExchangeOnlineManagement
  } 
  else 
  { 
   Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 }  

 if($OnlyOrganizationUsers.IsPresent)
 {
  #Check for Azure AD module
  $Module = Get-Module MsOnline -ListAvailable
  if($Module.count -eq 0) 
  { 
   Write-Host MSOnline module is not available  -ForegroundColor yellow  
   $Confirm= Read-Host Are you sure you want to install the module? [Y] Yes [N] No 
   if($Confirm -match "[yY]") 
   { 
    Write-host "Installing MSOnline PowerShell module"
    Install-Module MSOnline -Repository PSGallery -AllowClobber -Force
    Import-Module MSOnline
   } 
   else 
   { 
    Write-Host MSOnline module is required to generate the report.Please install module using Install-Module MSOnline cmdlet. 
    Exit
   }
  }
 }
}


 Install_Modules
 #Authentication using non-MFA
 if($NoMFA.IsPresent)
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
  if($OnlyOrganizationUsers.IsPresent)
  {
   Write-Host "Connecting Azure AD..."
   Connect-MsolService -Credential $Credential | Out-Null
  }
  Write-Host "Connecting Exchange Online PowerShell..."
  Connect-ExchangeOnline -Credential $Credential
 }
 #Connect to Exchange Online and AzureAD module using MFA 
 elseif($OnlyOrganizationUsers.IsPresent)
 {
  Write-Host "Connecting Exchange Online PowerShell..."
  Connect-ExchangeOnline
  Write-Host "Connecting Azure AD..."
  Connect-MsolService | Out-Null
 }
 #Connect to Exchange Online PowerShell
 else
 {
  Write-Host "Connecting Exchange Online PowerShell..."
  Connect-ExchangeOnline
 }


[DateTime]$MaxStartDate=((Get-Date).AddDays(-89)).Date 


<#Getting mail traffic data for past 90 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 [DateTime]$EndDate=(Get-Date).Date
 [DateTime]$StartDate=$MaxStartDate
}
#>

#Getting start date to audit mail traffic report
While($true)
{
 if ($StartDate -eq $null)
 {
  $StartDate=Read-Host Enter start time for report generation '(Eg:03/28/2022)'
 }
 Try
 {
  $Date=[DateTime]$StartDate
  if($Date -ge $MaxStartDate)
  { 
   break
  }
  else
  {
   Write-Host `nMail traffic can be retrieved only for past 90 days. Please select a date after $MaxStartDate -ForegroundColor Red
   return
  }
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}


#Getting end date to audit emails report
While($true)
{
 if ($EndDate -eq $null)
 {
  $EndDate=Read-Host Enter End time for report generation '(Eg: 03/28/2022)'
 }
 Try
 {
  $Date=[DateTime]$EndDate
  if($EndDate -lt ($StartDate))
  {
   Write-Host End time should be later than start time -ForegroundColor Red
   return
  }
  break
 }
 Catch
 {
  Write-Host `nNot a valid date -ForegroundColor Red
 }
}

$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
[DateTime]$CurrentStart=$StartDate
[DateTime]$CurrentEnd=$CurrentStart.AddMinutes(1439)


#Check whether CurrentEnd exceeds EndDate
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

if($CurrentStart -eq $CurrentEnd)
{
 Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
 Exit
}


if($MailsSent.isPresent)
{
 $Category="TopMailSender"
 $Header="Mails Sent"
  #Filter to get sender mail traffic for organization's users alone
  if($OnlyOrganizationUsers.IsPresent)
  {
   $Domains=(Get-MsolDomain).Name
  }
}

elseif($SpamsReceived.IsPresent)
{ 
 $Category="TopSpamRecipient"
 $Header="Spams Received"
}
elseif($MalwaresReceived.IsPresent)
{
 $Category="TopMalwareRecipient"
 $Header="Malwares Received"
}
else
{
 $Category="TopMailRecipient"
 $Header="Mails Received"
}


#Connect_Modules

Write-Host Getting mail traffic data... `n

#Output file declaration
$OutputCSV=".\Mail_Traffic_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
$AllMailTrafficData=@()
$AllMailTraffic=""
$AggregatedMailTrafficData=@()
$Page=1

While($True)
{
 Do
 { 
  Write-Progress -Activity "`n     Retrieving mail traffic data for $CurrentStart"
  $MailTrafficData=Get-MailTrafficSummaryReport -StartDate $CurrentStart -EndDate $CurrentEnd -Category $Category -Page $Page -PageSize 5000 | Select-Object C1,C2
  $AggregatedMailTrafficData+=$MailTrafficData
  $Page++
 }While($MailTrafficData.count -eq 5000)
 $Date=Get-Date($CurrentStart) -Format MM/dd/yyyy
 $DataCount=$AggregatedMailTrafficData.count
 for($i=0;$i -lt $DataCount ;$i++)
 {
  $UPN=$AggregatedMailTrafficData[$i].C1 
  if($OnlyOrganizationUsers.IsPresent)
  {
   $Domain=$UPN.Split("@") | Select-Object -Index 1
   if(($Domain -in $Domains) -eq $false)
   {
    continue
   }
  }
  $Count=$AggregatedMailTrafficData[$i].C2
  $AllMailTraffic=@{'Date'=$Date;'User Principal Name'=$UPN;$Header=$Count}
  $AllMailTrafficData= New-Object PSObject -Property $AllMailTraffic
  $AllMailTrafficData | select Date,'User Principal Name',$Header | Export-Csv $OutputCSV -NoTypeInformation -Append
 }
 $AggregatedMailTrafficData=@()
 $Page=1
 if([datetime]$CurrentEnd -ge [datetime]$EndDate)
 {
  break
 } 
 $CurrentStart=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
 $CurrentEnd=$CurrentEnd.AddMinutes($IntervalTimeInMinutes)
 if($CurrentStart -gt (Get-Date))
 {
  break
 }  
}

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue

if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host " The Output file available in the current working directory with name:" -NoNewline -ForegroundColor Yellow; 
  Write-Host $OutputCSV 
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$OutputCSV"   
  } 
 }