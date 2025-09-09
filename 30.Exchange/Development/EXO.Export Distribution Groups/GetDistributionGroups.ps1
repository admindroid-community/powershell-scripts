<#
=============================================================================================
Name:           Get distribution groups report
Version:        1.0
Website:        o365reports.com

Script Highlights:
~~~~~~~~~~~~~~~~~

1. Exports distribution groups with allowed / blocked external senders.
2. Helps to identify empty distribution groups. 
3. Allows to filter results based on group members count. 
4. Bulk import specific DLs and get the report on those specific DLs. 
5. The script automatically verifies and installs the Exchange Online PowerShell module (if not installed already) upon your confirmation.    
6. The script can be executed with MFA enabled account too.
7. It can be executed with Certificate-based Authentication (CBA) too. 
8. Exports DL’s members count based on Member Type such as user mailbox, group mailbox, etc.  
9. The script is scheduler friendly and exports the report result to CSV format. 
10. This script exports output in 2 CSV files. One with DL details and another with membership details. 


For detailed script execution:  https://o365reports.com/2025/03/04/export-distribution-groups-report-in-microsoft-365/ 
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$GroupNamesFile,
    [switch]$IsEmpty,
    [int]$MinGroupMembersCount,
    [Nullable[boolean]]$ExternalSendersBlocked = $null,
    [string]$UserName,
    [string]$Password,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint
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

  $Manager = @()
 if($_.ManagedBy.Count -gt 0)
 {
  foreach($ManageBy in $ManagedBy)
  { #Verify if manager property returned as ID and convert it to user-identifiable name
   if($ManageBy -match '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$')
   {
    if($ManagerHash.ContainsKey($ManageBy))
    {
     $ManagerName=$ManagerHash[$ManageBy]
    }
    # Retrieve the display name for the ID
    else
    {
     $ManagerName=(Get-EXORecipient -Identity $ManageBy).DisplayName
     $ManagerHash[$ManageBy]=$ManagerName
    }
    $ManageBy=$ManagerName
   }
   $Manager+=$ManageBy
  }
  $Manager= $Manager -join ","
 }
 $Recipient=""
 $RecipientHash=@{}
 for($KeyIndex = 0; $KeyIndex -lt $RecipientTypeArray.Length; $KeyIndex += 2)
 {
  $key=$RecipientTypeArray[$KeyIndex]
  $Value=$RecipientTypeArray[$KeyIndex+1]
  $RecipientHash.Add($key,$Value)
 }
 $Members=Get-DistributionGroupMember -ResultSize Unlimited -Identity $EmailAddress
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
   $MemberName=$Member.DisplayName

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
  $Result=@{'DisplayName'=$DisplayName;'PrimarySmtpAddress'=$EmailAddress;'Alias'=$Alias;'GroupType'=$GroupType;'Manager'=$Manager;'GroupMembersCount'=$MembersCount;'MembersCountByType'=$Recipient;'AuthorizedSenders'=$AuthorizedSenders;'ExternalSendersBlocked'=$ExternalSendersAllowed;'HiddenFromAddressList'=$_.HiddenFromAddressListsEnabled;
 'Description'=$_.Description;'CreationTime'=$_.WhenCreated;'DirSyncEnabled'=$_.IsDirSynced;'JoinGroupWithoutApproval'=$_.MemberJoinRestriction;'LeaveGroupWithoutApproval'=$_.MemberDepartRestriction} #Uncomment to print additional attributes in output
  $Results= New-Object PSObject -Property $Result 
  $Results | Select-Object DisplayName,PrimarySmtpAddress,Alias,GroupType,Manager,GroupMembersCount,AuthorizedSenders,ExternalSendersBlocked,MembersCountByType,HiddenFromAddressList,Description,CreationTime,DirSyncEnabled,JoinGroupWithoutApproval,LeaveGroupWithoutApproval | Export-Csv -Path $ExportSummaryCSV -Notype -Append
 }
}

#Print Detailed Output
Function Print_Output
{
 if($Print -eq 1)
 {
  $Result=@{'DisplayName'=$DisplayName;'PrimarySmtpAddress'=$EmailAddress;'Alias'=$Alias;'Members'=$MemberName;'MemberEmail'=$MemberEmail;'MemberType'=$RecipientTypeDetail} 
  $Results= New-Object PSObject -Property $Result 
  $Results | Select-Object DisplayName,PrimarySmtpAddress,Alias,Members,MemberEmail,MemberType | Export-Csv -Path $ExportCSV -Notype -Append
 }
}


Function main()
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

 
 #Set output file 
 $ExportCSV=".\DistributionGroup-MembershipReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" #Detailed report
 $ExportSummaryCSV=".\DistributionGroup-Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" #Summary report

 #Get a list of RecipientTypeDetail
 $RecipientTypeArray=Get-Content -Path .\RecipientTypeDetails.txt -ErrorAction Stop
 $Result=""  
 $Results=@()
 $ManagerHash = @{}
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
  Write-Host `n" Detailed report available in:" -NoNewline -ForegroundColor Yellow
  Write-Host $ExportCSV `n
  Write-host " Summary report available in:" -NoNewline -ForegroundColor Yellow
  Write-Host $ExportSummaryCSV 
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n 
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
 #Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
 
}
 . main