#Accept input parameter
Param
(
    [Parameter(Mandatory = $false)]
    [string]$GroupNamesFile,
    [switch]$IsEmpty,
    [int]$MinGroupMembersCount,
    [string]$UserName,
    [string]$Password,
    [Switch]$MFA
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
  $Result=@{'DisplayName'=$DisplayName;'PrimarySmtpAddress'=$EmailAddress;'Alias'=$Alias;'Manager'=$Manager;'GroupMembersCount'=$Members.Count; 'Members'=$Member;'MemberType'=$RecipientTypeDetail} 
  $Results= New-Object PSObject -Property $Result 
  $Results | Select-Object DisplayName,PrimarySmtpAddress,Alias,Manager,GroupMembersCount,Members,MemberType | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

Function main()
{
 #Connect AzureAD and Exchange Online from PowerShell 
 Get-PSSession | Remove-PSSession 

 #Get a list of RecipientTypeDetail
 $RecipientTypeArray=Get-Content -Path .\RecipientTypeDetails.txt -ErrorAction Stop

 #Authentication using MFA
 if($MFA.IsPresent)
 {
  $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
  If ($MFAExchangeModule -eq $null)
  {
   Write-Host  `nPlease install Exchange Online MFA Module.  -ForegroundColor yellow
   
   Write-Host You can install module using below blog : `nLink `nOR you can install module directly by entering "Y"`n
   $Confirm= Read-Host Are you sure you want to install module directly? [Y] Yes [N] No
   if($Confirm -match "[yY]")
   {
     Write-Host Yes
     Start-Process "iexplore.exe" "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"
   }
   else
   {
    Start-Process 'https://http://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
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
     Start-Process 'https:http://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
     Exit
    }
     
   }
  
  #Importing Exchange MFA Module
  . "$MFAExchangeModule"
  Write-Host Enter credential in prompt to connect to Exchange Online
  Connect-EXOPSSession -WarningAction SilentlyContinue
  Write-Host Connected to Exchange Online
  Write-Host `nReport generation in progress...
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
  Connect-MsolService -Credential $credential 
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
  Import-PSSession $Session -CommandName Get-DynamicDistributionGroup,Get-Recipient -FormatTypeName * -AllowClobber | Out-Null
 }







 

 #Set output file 
 $ExportCSV=".\DynamicDistributionGroup-DetailedMembersReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" #Detailed report
 $ExportSummaryCSV=".\DynamicDistributionGroup-SummaryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" #Summary report

 
 $Result=""  
 $Results=@() 
 $Count=0

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
 #Clean up session 
 Get-PSSession | Remove-PSSession
}
. main
