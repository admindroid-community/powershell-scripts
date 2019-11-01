Param
(
    [Parameter(Mandatory = $false)]
    [string]$GroupNamesFile,
    [switch]$IsEmpty,
    [int]$MinGroupMembersCount,
    [switch]$MFA,
    [Nullable[boolean]]$ExternalSendersBlocked = $null,
    [string]$UserName,
    [string]$Password
)

Function Get_members
{
 $DisplayName=$_.DisplayName
 Write-Progress -Activity "`n     Processed Group count: $Count "`n"  Getting members of: $DisplayName"
 $Alias=$_.Alias
 $EmailAddress=$_.PrimarySmtpAddress
 $GroupType=$_.GroupType
 $ManagedBy=$_.ManagedBy
 $ExternalSendersAllowed=$_.RequireSenderAuthenticationEnabled
 if(($ExternalSendersBlocked -ne $null) -and ($ExternalSendersBlocked -ne $ExternalSendersAllowed))
 {
  $Print=0
 }
 #Get Distribution Group Authorized Senders
 $AcceptMessagesOnlyFrom=$_.AcceptMessagesOnlyFromSendersOrMembers
 $AuthorizedSenders=""
 if($AcceptMessagesOnlyFrom.Count -gt 0)
 {
  foreach($item in $AcceptMessagesOnlyFrom)
  {
   $AuthorizedSenders=$AuthorizedSenders+$item
   if($AcceptMessagesOnlyFrom.indexof($item) -lt (($AcceptMessagesOnlyFrom.count)-1))
   {
    $AuthorizedSenders=$AuthorizedSenders+","
   }
  }
 }
 elseif($ExternalSendersAllowed -eq "True")
 {
  $AuthorizedSenders="Only Senders in Your Organization"
 }
 else
 {
  $AuthorizedSenders="Senders inside & Outside of Your Organization"
 }

 $Manager=""
 if($_.ManagedBy.Count -gt 0)
 {
  foreach($ManageBy in $ManagedBy)
  {
   $Manager=$Manager+$ManageBy
   if($ManagedBy.indexof($ManageBy) -lt (($ManagedBy.count)-1))
   {
    $Manager=$Manager+","
   }
  }
 }
 $Recipient=""
 $RecipientHash=@{}
 for($KeyIndex = 0; $KeyIndex -lt $RecipientTypeArray.Length; $KeyIndex += 2)
 {
  $key=$RecipientTypeArray[$KeyIndex]
  $Value=$RecipientTypeArray[$KeyIndex+1]
  $RecipientHash.Add($key,$Value)
 }
 $Members=Get-DistributionGroupMember -ResultSize Unlimited -Identity $DisplayName
 $MembersCount=($Members.name).Count

 #GroupSize Filter
 if(([int]$MinGroupMembersCount -ne "") -and ($MembersCount -lt [int]$MinGroupMembersCount))
 {
  $Print=0
 }

 #Check for Empty Group
 elseif($MembersCount -eq 0)
 {
  $Member="No Members"
  $MemberEmail="-"
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
   $RecipientTypeDetail=$Member.RecipientTypeDetails
   $MemberEmail=$Member.PrimarySMTPAddress
   if($MemberEmail -eq "")
   {
    $MemberEmail="-"
   }
   #Get Counts by RecipientTypeDetail
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
 
 #Print Summary report
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
  $Result=@{'DisplayName'=$DisplayName;'PrimarySmtpAddress'=$EmailAddress;'Alias'=$Alias;'GroupType'=$GroupType;'Manager'=$Manager;'GroupMembersCount'=$MembersCount;'MembersCountByType'=$Recipient;'AuthorizedSenders'=$AuthorizedSenders;'ExternalSendersBlocked'=$ExternalSendersAllowed <#;'HiddenFromAddressList'=$_.HiddenFromAddressListsEnabled;
 'Description'=$_.Description;'CreationTime'=$_.WhenCreated;'DirSyncEnabled'=$_.IsDirSynced;'JoinGroupWithoutApproval'=$_.MemberJoinRestriction;'LeaveGroupWithoutApproval'=$_.MemberDepartRestriction #>} #Uncomment to print additional attributes in output
  $Results= New-Object PSObject -Property $Result 
  $Results | Select-Object DisplayName,PrimarySmtpAddress,Alias,GroupType,Manager,GroupMembersCount,AuthorizedSenders,ExternalSendersBlocked,MembersCountByType <#,HiddenFromAddressList,Description,CreationTime,DirSyncEnabled,JoinGroupWithoutApproval,LeaveGroupWithoutApproval #>  | Export-Csv -Path $ExportSummaryCSV -Notype -Append
 }
}

#Print Detailed Output
Function Print_Output
{
 if($Print -eq 1)
 {
  $Result=@{'DisplayName'=$DisplayName;'PrimarySmtpAddress'=$EmailAddress;'Alias'=$Alias;'Members'=$Member;'MemberEmail'=$MemberEmail;'MemberType'=$RecipientTypeDetail} 
  $Results= New-Object PSObject -Property $Result 
  $Results | Select-Object DisplayName,PrimarySmtpAddress,Alias,Members,MemberEmail,MemberType | Export-Csv -Path $ExportCSV -Notype -Append
 }
}


Function main()
{
 #Clean up session 
 Get-PSSession | Remove-PSSession

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
     Start-Process 'https://http://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/'
     Exit
    }    
   }
 
  #Importing Exchange MFA Module
  . "$MFAExchangeModule"
  Write-Host Enter credential in prompt to connect to Exchange Online
  Connect-EXOPSSession -WarningAction SilentlyContinue
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
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
  Import-PSSession $Session -CommandName Get-DistributionGroup,Get-DistributionGroupMember -FormatTypeName * -AllowClobber | Out-Null
 }
 
 #Set output file 
 $ExportCSV=".\DistributionGroup-DetailedMembersReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" #Detailed report
 $ExportSummaryCSV=".\DistributionGroup-SummaryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" #Summary report

 #Get a list of RecipientTypeDetail
 $RecipientTypeArray=Get-Content -Path .\RecipientTypeDetails.txt -ErrorAction Stop
 $Result=""  
 $Results=@()

 #Check for input file
 if([string]$GroupNamesFile -ne "") 
 { 
  #We have an input file, read it into memory 
  $DG=@()
  $DG=Import-Csv -Header "DisplayName" $GroupNamesFile
  foreach($item in $DG)
  {
   Get-DistributionGroup -Identity $item.displayname | Foreach{
   $Print=1
   Get_Members}
   $Count++
  }
 }
 else
 {
  #Get all distribution group
  Get-DistributionGroup -ResultSize Unlimited | Foreach{
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
  Write-Host No DistributionGroup found
 }
 #Clean up session 
 Get-PSSession | Remove-PSSession
 
}
 . main