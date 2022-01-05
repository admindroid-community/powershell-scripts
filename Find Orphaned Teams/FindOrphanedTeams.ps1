#Accept input paramenters 
param(
[string]$UserName, 
[string]$Password
)

Function PrintOutput
{
 $Result = @{'Team Name'=$Displayname;'Description'=$Description;'Member Count'=$MemberCount;'Guest Count'=$GuestCount;'Team Type'=$Visibility;'IsArchived'=$IsArchived ;'Group Id'=$Groupid} 
 $Results = New-Object PSObject -Property $Result 
 $Results |select-object 'Team Name','Description','Member Count','Guest Count','Team Type','IsArchived','Group Id' | Export-CSV $ExportCSV  -NoTypeInformation -Append
}


#Connect to Microsoft Teams
$Module=Get-Module -Name MicrosoftTeams -ListAvailable 
if($Module.count -eq 0)
{
 Write-Host MicrosoftTeams module is not available  -ForegroundColor yellow 
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
 if($Confirm -match "[yY]")
 {
  Write-Host Installing Microsoft Teams PowerShell module... -ForegroundColor Yellow
  Install-Module MicrosoftTeams -Repository PSGalleryInt
 }
 else
 {
  Write-Host MicrosoftTeams module is required.Please install the module using Install-Module MicrosoftTeams cmdlet.
  Exit
 }
}

Write-Host Connecting to Microsoft Teams... -ForegroundColor Yellow
#For scheduling purpose passing credential as parameter
if(($UserName -ne "") -and ($Password -ne ""))
{
 $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
 $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
 $Team=Connect-MicrosoftTeams -Credential $Credential
}
else
{  
 $Team=Connect-MicrosoftTeams
}
$ExportCSV=".\TeamsWithoutOwnersReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$Result="" 
$Results=@() 
$ExportedCount=0
$Count=0


#Retrieve all the Teams
Write-Host Retrieving Teams details...
Get-Team | foreach {
 $DisplayName=$_.DisplayName
 $Count++
 Write-Progress -Activity "Processed Teams Count: $Count" "Currently Processing Team: $DisplayName" 
 $Description=$_.Description
 $Visibility=$_.Visibility
 $IsArchived=$_.Archived
 $GroupID=$_.Groupid
 $MemberCount=0
 $GuestCount=0
 $OwnerCount=0
 Get-TeamUser -GroupId $GroupID | foreach {
  if($_.role -eq "owner")
  { $OwnerCount++ }
  elseif($_.role -eq "member")
  { $MemberCount++ }
  else
  { $GuestCount++ }
 }
 if($OwnerCount -eq 0)
 {
  $ExportedCount++
  PrintOutput
 }
} 

#Open output file after execution
If($ExportedCount -eq 0)
{
 
 Write-Host No records found
}
else
{
 Write-Host `nThe output file contains $ExportedCount orphaned Teams.
 if((Test-Path -Path $ExportCSV) -eq "True") 
 {
  Write-Host `nThe Output file availble in $ExportCSV -ForegroundColor Green
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$ExportCSV"   
  } 
 }
}
<#
=============================================================================================
Name:           Find orphaned teams in Microsoft Teams
website:        o365reports.com
For detailed Script execution: https://o365reports.com/2022/01/05/finding-and-managing-microsoft-teams-without-owner
============================================================================================
#>