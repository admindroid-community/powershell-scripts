Param
(
    [Parameter(Mandatory = $false)]
    [switch]$MFA,
    [switch]$OnlyOrganizationUsers,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserName,
    [string]$Password
)

#Print Output
function Print_Output
{
 $AllMailTraffic=@{'Date'=$PreviousDate;'User Principal Name'=$PreviousName;'Mails Sent'=$MailsSent;'Mails Received'=$MailsReceived}
 $AllMailTrafficData= New-Object PSObject -Property $AllMailTraffic
 $AllMailTrafficData | select Date,'User Principal Name','Mails Sent','Mails Received' | Export-Csv $OutputCSV -NoTypeInformation -Append
 $Print=1
}

#Get inbound and outbound mail count
function Get_MailCount
{ 
 if($MailTraffic.Direction -eq "Inbound")
 {
  $Global:MailsReceived += $MailTraffic.MessageCount
 }
 else
 {
  $Global:MailsSent += $MailTraffic.MessageCount
 }
 $Print=0
}

function main{


 #Getting StartDate and EndDate for mail traffic collection
 if ((($StartDate -eq $null) -and ($EndDate -ne $null)) -or (($StartDate -ne $null) -and ($EndDate -eq $null)))
 {
  Write-Host `nPlease enter both StartDate and EndDate for mail traffic data collection -ForegroundColor Red
  exit
 }
 elseif(($StartDate -eq $null) -and ($EndDate -eq $null))
 {
  $StartDate=(((Get-Date).AddDays(-90))).Date
  $EndDate=Get-Date
 }
 else
 {
  $StartDate=[DateTime]$StartDate
  $EndDate=[DateTime]$EndDate
  if($StartDate -lt ((Get-Date).AddDays(-90)))
  {
   Write-Host `nMail traffic data can be retrieved only for past 90 days. Please select a date after (Get-Date).AddDays(-90) -ForegroundColor Red
   Exit
  }
  if($EndDate -lt ($StartDate))
  {
   Write-Host `nEnd time should be later than start time -ForegroundColor Red
   Exit
  }
 }

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
  } 
  else 
  { 
   Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
   Exit
  }
 } 

 #Connect Exchange Online with MFA
 if($MFA.IsPresent)
 {
  Connect-ExchangeOnline
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
  Connect-ExchangeOnline -Credential $Credential
 }

 #Connect to MsolService to get internal domains
 if($OnlyOrganizationUsers.IsPresent)
 {
  #Connect MsolService
  if($mfa.IsPresent)
  {
   Connect-MsolService
  }
  else
  {
   Connect-MsolService -Credential $Credential
  }
  $Domains=(Get-MsolDomain).Name
 }

 #Output file declaration
 $OutputCSV=".\Mail_Traffic_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
 $AllMailTrafficData=@()
 $AllMailTraffic=""
 $AggregatedMailTrafficData=@()
 $Page=1
 Write-Host Getting mail traffic data...
 Do
 {
  $MailTrafficData=Get-MailTrafficTopReport -StartDate $StartDate -EndDate $EndDate -Page $Page -PageSize 5000 | Sort-Object Date,Name,Direction
  $AggregatedMailTrafficData+=$MailTrafficData
  $Page++
 }While($MailTrafficData.count -eq 5000)

 Write-Host The script has found $AggregatedMailTrafficData.count records
 Write-Host Processing Mail traffic data...
 $PreviousDate=$MailTrafficData[0].Date #FirstRecordDate
 $PreviousName=$MailTrafficData[0].Name #FirstRecordName
 $Global:MailsSent=0
 $Global:MailsReceived=0
 $Print=1
 $ProcessedRecords=0
 foreach($MailTraffic in $AggregatedMailTrafficData)
 {
  $ProcessedRecords++
  Write-Progress -Activity "`n     Processing Mail traffic data from $StartDate to $EndDate.."`n" Processed record count: $ProcessedRecords"
  $CurrentDate=$MailTraffic.Date
  $CurrentName=$MailTraffic.Name
  $Count=$MailTraffic.MessageCount

  #Filter to get mail traffic for organization's users alone
  if($OnlyOrganizationUsers.IsPresent)
  {
   $Domain=$CurrentName.Split("@") | Select-Object -Index 1
   if(($Domain -in $Domains) -eq $false)
   {
    continue
   }
  }
  if($PreviousDate -eq $CurrentDate)
  {
   if($PreviousName -eq $CurrentName)
   {
    Get_MailCount -traffic $MailTraffic
   }
   else
   {
    #Export Output to CSV
    Print_Output
    $PreviousName=$CurrentName
    $Global:MailsReceived=0
    $Global:MailsSent=0
    Get_MailCount -Traffic $MailTraffic
   }
  }
  else
  {
   If($Print -eq 0)
   {
    #Export Output to CSV
    Print_Output
   }
   $PreviousDate=$CurrentDate
   $PreviousName=$CurrentName
   $Global:MailsReceived=0
   $Global:MailsSent=0
   Get_MailCount -Traffic $MailTraffic
  }
 }
 #Export Output to CSV
 Print_Output

 If($ProcessedRecords -eq 0)
 {
  Write-Host No records found
 }
 #Open output file after execution
 else
 {
  if((Test-Path -Path $OutputCSV) -eq "True")
  {
   Write-Host `nThe Output file available in $OutputCSV -ForegroundColor Green
   $Prompt = New-Object -ComObject wscript.shell
   $UserInput = $Prompt.popup("Do you want to open output file?",`
 0,"Open Output File",4)
   If ($UserInput -eq 6)
   {
    Invoke-Item "$OutputCSV"
   }
  }
 }
}

 . main