<#
=============================================================================================
Name:           Auto Archive Teams in Microsoft Teams
Version:        1.0
Website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Export all inactive teams along with their last activity date and inactive days. 
2. Automatically archives inactive teams based on their inactivity period.
3. List all teams that have had no activity since creation. 
4. Find teams that have been inactive for a specific period.
5. Automatically install the Microsoft Graph PowerShell module (if not installed already) upon your confirmation.
6. The script can be executed with an MFA-enabled account too.
7. Supports Certificate-based Authentication too.
8. The script is scheduler friendly.Sends password expiry notifications to users about upcoming password expiry.


For detailed Script execution: https://o365reports.com/2025/03/18/how-to-archive-inactive-teams-in-microsoft-teams/
============================================================================================
#>

Param
(
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [int]$InactiveDays,
    [Switch]$IncludeTeamsWithNoActivity,
    [Switch]$ArchiveInactiveTeams,
    [switch]$Force
    
)
Function Connect_MgGraph
{
 #Check for module installation
 $Module=Get-Module -Name Microsoft.Graph -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Microsoft Graph PowerShell SDK is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Microsoft Graph PowerShell module..."
   Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
  }
  else
  {
   Write-Host "Microsoft Graph PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph cmdlet." 
   Exit
  }
 }
 #Disconnect Existing MgGraph session
 if($CreateSession.IsPresent)
 {
  Disconnect-MgGraph | Out-Null
 }


 Write-Host Connecting to Microsoft Graph...
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
 }
 else
 {
  Connect-MgGraph -Scopes "Reports.Read.All","TeamSettings.ReadWrite.All"  -NoWelcome
 }
}
Connect_MgGraph


$Location=Get-Location
$TempFile="$Location\TeamsUsageReport_tempFile_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv" 
$ExportCSV = "$Location\ArchiveInactiveTeams_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$Count=0
$PrintedTeams=0
$ArchiveStatus="-"

if($ArchiveInactiveTeams.IsPresent)
{
 if($InactiveDays -eq "")
 {
  Write-Host `nInactive days is mandatory to archive inactive teams. -ForegroundColor Magenta
  $InactiveDays= Read-host Enter Inactive days
 }
 #Getting consent for archiving inactive teams
 if(!($Force.IsPresent))
 {
  $Confirm= Read-Host `nDo you want to archive teams that are inactive for $InactiveDays days [Y] Yes [N] No 
  if($Confirm -notmatch "[yY]") 
  { 
   Write-host Exiting script...
   Exit
  }
 }
}

#Retrieving Teams usage report
Try
{
 Get-MgReportTeamActivityDetail -Period 'D7' -OutFile $TempFile
}
Catch
{
Write-Host Unable to fetch teams usage report. Error occurred - $($Error[0].Exception.Message) -ForegroundColor Red
Exit
}

Write-Host Generating inactive teams report...
#Import and process teams usage report
Import-Csv -Path $TempFile | foreach {
 $TeamName=$_.'Team Name'
 $Count++
 $Print=1
 Write-Progress -Activity "`n     Processed teams: $Count - $EnterpriseAppName "
 $TeamType=$_.'Team Type'
 $LastActivityDate=$_.'Last Activity Date'
 if($LastActivityDate -eq "")
 {
  $LastActivityDate = "Never Active"
  $InactivePeriod = "-"
 }
 else
 {
  $InactivePeriod=(New-TimeSpan -Start $LastActivityDate).Days
 }
 $IsDeleted=$_.'Is Deleted'
 $TeamId=$_.'Team Id'

 #Filter teams based on inactive days
 if($InactivePeriod -ne "-")
 {
  if(($InactiveDays -ne "") -and ($InactiveDays -gt $InactivePeriod))
  {
   $Print=0
  }
 }

 #Filter for excluding never active teams
 if(!($IncludeTeamsWithNoActivity.IsPresent) -and ($LastActivityDate -eq "Never Active"))
 {
  $Print=0
 }

 #Exclude deleted teams
 if($IsDeleted -eq $true)
 {
  $Print=0
 }

 if(($ArchiveInactiveTeams.IsPresent) -and ($Print -eq 1))
 {
  #Check if team is already archived
  if((Get-mgteam -TeamId $TeamId).IsArchived)
  {
   $ArchiveStatus="Team is already archived"
  }
  else
  {
   Invoke-MgArchiveTeam -TeamId $TeamId -ShouldSetSpoSiteReadOnlyForMembers
   if($?)
   {
    $ArchiveStatus="Successfully archived"
   }
   else
   {
    $ArchiveStatus="Error occurred"
   }
  }
 }


 #Export the output to CSV
 if($Print -eq 1)
 {
  $PrintedTeams++
  $ExportResult=[PSCustomObject]@{'Team Name'=$TeamName;'Team Type'=$TeamType;'Last Activity Date'=$LastActivityDate;'Inactive Days'=$InactivePeriod;'Archive Status Log'=$ArchiveStatus}
  $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

#Remove/delete the reference file
Remove-Item $TempFile
Disconnect-MgGraph | Out-Null
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
  
#Open Output file after execution
 if((Test-Path -Path $ExportCSV) -eq "True") 
 {
  Write-Host `The exported report contains $PrintedTeams teams .
  Write-Host `nDetailed report available in: -NoNewline -Foregroundcolor Yellow; Write-Host $ExportCSV
   $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$ExportCSV"   
  } 
 }
 else
 {
  Write-Host No teams found for the given criteria.
 }
