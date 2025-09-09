<#
=============================================================================================
Name:           Find orphaned teams in Microsoft Teams
Description:    This script exports teams without owner to CSV file
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script can be executed with MFA enabled account too. 
2. This script uses modern authentication to connect to Microsoft Teams PowerShell.
3. Automatically installs Microsoft Teams module (if not installed already) upon your confirmation. 
4. Lastly, it exports orphaned teams to CSV file.
5. Also, the script is scheduler-friendly. I.e., Credential can be passed as a parameter instead of saving inside the script. 

For detailed Script execution: https://o365reports.com/2022/01/05/finding-and-managing-microsoft-teams-without-owner
============================================================================================
#>

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
  Write-Host `n" The Output file availble in:" -NoNewline -ForegroundColor Yellow; Write-Host $ExportCSV 
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline;
  Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$ExportCSV"   
  } 
 }
}