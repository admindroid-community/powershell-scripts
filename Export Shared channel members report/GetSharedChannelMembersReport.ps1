<#
=============================================================================================
Name:           Microsoft Teams Shared Channel Members Report
website:        o365reports.com
For detailed Script execution:  https://o365reports.com/2023/02/28/ms-teams-shared-channel-membership-report
============================================================================================
#>
param(
[string]$UserName, 
[string]$Password, 
[string]$CertificateThumbprint,
[string]$ApplicationId,
[string]$TenantId,
[string]$TeamName
) 

#Connect to Microsoft Teams
$Module=Get-Module -Name MicrosoftTeams -ListAvailable 
if($Module.count -eq 0)
{
 Write-Host MicrosoftTeams module is not available  -ForegroundColor yellow 
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
 if($Confirm -match "[yY]")
 {
  Install-Module MicrosoftTeams -Scope CurrentUser
 }
 else
 {
  Write-Host MicrosoftTeams module is required.Please install module using Install-Module MicrosoftTeams cmdlet.
  Exit
 }
}

#Storing credential in script for scheduling purpose/ Passing credential as parameter
if(($UserName -ne "") -and ($Password -ne ""))
{
 $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
 $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
 Connect-MicrosoftTeams -Credential $Credential
}
elseif(($CertificateThumbprint -ne "") -and ($ApplicationId -ne "") -and ($TenantId -ne ""))
{
 Connect-MicrosoftTeams -CertificateThumbprint $CertificateThumbprint -ApplicationId $ApplicationId -TenantId $TenantId
}
else
{  
 Connect-MicrosoftTeams | Out-Null
}

Function Process_TeamChannel
{
 Get-TeamAllChannel -GroupId $TeamId -MembershipType Shared | foreach {
  $global:SharedChannelCount++
  $ChannelName=$_.DisplayName
  $Description=$_.Description
  $HostTeamId=$_.HostTeamId
  if($Description -eq $null)
  {
   $Description="-"
  }
  if($HostTeamId -eq $TeamId )
  {
   $SharedChannelType="Team hosted channel"
  }
  else
  {
   $SharedChannelType="Incoming channel"
  }
  $Membership=Get-TeamChannelUser -GroupId $HostTeamId -DisplayName $ChannelName
  $AllMembers= (@($Membership.user)-join ',')
  $MembersCount=$Membership.Count

  #Exporting shared channel summary report
  $Result = @{'Team Name'=$Teamname;'Shared Channel Name'=$ChannelName;'Shared Channel Type'=$SharedChannelType;'Description'=$Description;'Channel Members'=$AllMembers;'Members Count'=$MembersCount }
  $Results = New-Object PSObject -Property $Result
  $Results |select-object 'Team Name','Shared Channel Name','Shared Channel Type','Description','Members Count','Channel Members' | Export-Csv -Path $OutputCSVName -Notype -Append

  foreach($Member in $Membership)
  {
   $Name=$Member.Name
   $Email=$Member.User
   $Role=$Member.Role
   $Output=@{'Team Name'=$Teamname;'Shared Channel Name'=$ChannelName;'Shared Channel Type'=$SharedChannelType;'Description'=$Description;'Member Name'=$Name;'Member Email'=$Email;'Role'=$Role}
   $Outputs = New-Object PSObject -Property $Output
   $Outputs | Select-Object 'Team Name','Shared Channel Name','Shared Channel Type','Description','Member Name','Member Email','Role' | Export-Csv -Path $OutputCSV -Notype -Append
  }
 }
}

$Result=""  
$Results=@()
$Output=""
$Outputs=@() 

$OutputCSVName=".\SharedChannelReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
$OutputCSV=".\Detailed_SharedChannelMembershipReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
Write-Host Generating shared channel membership report...  -ForegroundColor Yellow
$ProcessedCount=0
$global:SharedChannelCount=0

#Process all teams
if($TeamName -eq "")
{
 Get-Team | foreach {
  $TeamName=$_.DisplayName
  $TeamId=$_.GroupId
  $ProcessedCount++
  Write-Progress -Activity "`n     Processed teams count: $ProcessedCount "`n"  Currently Processing team: $TeamName"
  Process_TeamChannel
 }
}

else
{
 $TeamDetails=Get-Team -DisplayName $TeamName     
 if($TeamDetails -eq $null)
 {
  Write-Host Team $TeamName is not available. Please check the team name. -ForegroundColor Red
 }
 else
 {
  $TeamId=$TeamDetails.GroupId  
  Process_TeamChannel
 }
}

#Open output file after execution  
Write-Host `nScript executed successfully 
Write-Host `n$SharedChannelCount shared channels exported in the output file.
if((Test-Path -Path $OutputCSVName) -eq "True") 
{
 Write-Host "Shared channel summary report available in: $OutputCSVName"  -ForegroundColor Green 
 Write-Host "Detailed shared channel membership report available in: $OutputCSV"  -ForegroundColor Green
 $Prompt = New-Object -ComObject wscript.shell   
 $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
 If ($UserInput -eq 6)   
 {   
  Invoke-Item "$OutputCSVName"
  Invoke-Item "$OutputCSV"   
 }  
} 