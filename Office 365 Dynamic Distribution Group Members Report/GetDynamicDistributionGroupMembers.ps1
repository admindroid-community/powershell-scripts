<#
=============================================================================================
Name:           Export Dynamic Distribution Group Members Report
Version:        2.0
Website:        o365reports.com
Script by:      O365Reports Team
For detailed script execution:  https://o365reports.com/2019/03/23/export-dynamic-distribution-group-members-to-csv/
============================================================================================
#>

#Accept input parameter
Param
(
    [Parameter(Mandatory = $false)]
    [string]$GroupNamesFile,
    [switch]$IsEmpty,
    [int]$MinGroupMembersCount,
    [string]$UserName,
    [string]$Password,
    [Switch]$NoMFA
)

#Get group members
Function Get_Members
{
 $DisplayName=$_.DisplayName
 Write-Progress -Activity "`n     Processed Group count: $Count "`n"  Getting members of: $DisplayName"
 $Alias=$_.Alias
 $EmailAddress=$_.PrimarySmtpAddress
 $HiddenFromAddressList=$_.HiddenFromAddressListsEnabled
 $RecipientFilter=$_.RecipientFilter
 $RecipientHash=@{}
 for($KeyIndex = 0; $KeyIndex -lt $RecipientTypeArray.Length; $KeyIndex += 2)
 {
  $key=$RecipientTypeArray[$KeyIndex]
  $Value=$RecipientTypeArray[$KeyIndex+1]
  $RecipientHash.Add($key,$Value)
 }
 $Manager=$_.ManagedBy
 if($Manager -eq $null)
 { 
  $Manager="-"
 }
 $Recipient=""
 $Members=Get-Recipient -ResultSize unlimited -RecipientPreviewFilter $RecipientFilter
 #GroupSize Filter
 if(([int]$MinGroupMembersCount -ne "") -and ($Members.count -lt [int]$MinGroupMembersCount))
 {
  $Print=0
 }

 #Empty Group Filter
 elseif($Members.Count -eq 0)
 {
  $Member="No Members"
  $RecipientTypeDetail="-"
  Print_Output
 }

 #Loop through each member in a group
 else
 {
  foreach($Member in $Members)
  {
   if($IsEmpty.IsPresent)
   {
    $Print=0
    break
   }
   #Get Counts by RecipientTypeDetail
   $RecipientTypeDetail=$Member.RecipientTypeDetails
   $MemberEmail=$Member.PrimarySMTPAddress
   foreach($key in [object[]]$Recipienthash.Keys)
   {
    if(($RecipientTypeDetail -eq $key) -eq "true")
    {
     [int]$RecipientHash[$key]+=1
    }
   }
   Print_Output
  }
 }
 
 #Export Summary Report
 if($Print -eq 1)
 {
  #Order RecipientTypeDetail based on count
  $Hash=@{}
  $Hash=$RecipientHash.GetEnumerator() | Sort-Object -Property value -Descending |foreach{
   if([int]$($_.Value) -gt 0 )
   {
    if($Recipient -ne "")
    { $Recipient+=";"} 
    $Recipient+=@("$($_.Key) - $($_.Value)")    
   }
   if($Recipient -eq "")
   {$Recipient="-"}
  }
  $Output=@{'DisplayName'=$DisplayName;'PrimarySmtpAddress'=$EmailAddress;'Alias'=$Alias;'Manager'=$Manager;'GroupMembersCount'=$Members.Count;'HiddenFromAddressList'=$HiddenFromAddressList;'MembersCountByType'=$Recipient} 
  $Outputs= New-Object PSObject -Property $Output
  $Outputs | Select-Object DisplayName,PrimarySmtpAddress,Alias,Manager,HiddenFromAddressList,GroupMembersCount,MembersCountByType | Export-Csv -Path $ExportSummaryCSV -Notype -Append
 }
}

#Print Detailed Output
Function Print_Output
{
 if($Print=1)
 {
  $Result=@{'DisplayName'=$DisplayName;'PrimarySmtpAddress'=$EmailAddress;'Alias'=$Alias;'Manager'=$Manager;'GroupMembersCount'=$Members.Count; 'Members'=$Member; 'MemberEmail'= $MemberEmail; 'MemberType'=$RecipientTypeDetail} 
  $Results= New-Object PSObject -Property $Result 
  $Results | Select-Object DisplayName,PrimarySmtpAddress,Alias,Manager,GroupMembersCount,Members,MemberEmail,MemberType | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

Function main()
{

 #Get a list of RecipientTypeDetail
 $RecipientTypeArray=Get-Content -Path .\RecipientTypeDetails.txt -ErrorAction Stop

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
  Write-Host "Connecting Azure AD..."
  Connect-MsolService -Credential $Credential | Out-Null
  Write-Host "Connecting Exchange Online PowerShell..."
  Connect-ExchangeOnline -Credential $Credential
 }
 #Connect to Exchange Online and AzureAD module using MFA 
 else
 {
  Write-Host "Connecting Exchange Online PowerShell..."
  Connect-ExchangeOnline
  Write-Host "Connecting Azure AD..."
  Connect-MsolService | Out-Null
 }

 #Friendly DateTime conversion
 if($friendlyTime.IsPresent)
 {
  If(((Get-Module -Name PowerShellHumanizer -ListAvailable).Count) -eq 0)
  {
   Write-Host Installing PowerShellHumanizer for Friendly DateTime conversion
   Install-Module -Name PowerShellHumanizer
  }
 }

 #Set output file 
 $ExportCSV=".\DynamicDistributionGroup-DetailedMembersReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" #Detailed report
 $ExportSummaryCSV=".\DynamicDistributionGroup-SummaryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" #Summary report

 
 $Result=""  
 $Results=@() 
 $Count=1

 #Check for input file
 if ($GroupNamesFile -ne "") 
 { 
  #We have an input file, read it into memory 
  $DDG=@()
  $DDG=Import-Csv -Header "DisplayName" $GroupNamesFile
  foreach($item in $DDG)
  {
   Get-DynamicDistributionGroup -Identity $item.displayname | Foreach{
   $Print=1
   Get_Members}
   $Count++
  }
 }
 else
 {
  #Get all dynamic distribution group
  Get-DynamicDistributionGroup | Foreach{
  $Print=1
  Get_Members
  $Count++}
 }

 #Open output file after execution 
 Write-Host `nScript executed successfully
 if((Test-Path -Path $ExportCSV) -eq "True")
 {
  Write-Host Detailed report available in: $ExportCSV
  Write-host Summary report available in: $ExportSummaryCSV
  $Prompt = New-Object -ComObject wscript.shell  
  $UserInput = $Prompt.popup("Do you want to open output file?",`  
  0,"Open Output File",4)  
  If ($UserInput -eq 6)  
  {  
   Invoke-Item "$ExportCSV"  
   Invoke-Item "$ExportSummaryCSV"
  } 
 }
 Else
 {
  Write-Host No DynamicDistributionGroup found
 }
 Write-Host "For more Office 365 reports, do check AdminDroid Office 365 reporting tool." -ForegroundColor Cyan
 #Clean up session 
 Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
}
. main
