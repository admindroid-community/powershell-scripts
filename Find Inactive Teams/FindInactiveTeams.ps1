<#
=============================================================================================
Name:           Find All Inactive Teams in Microsoft Teams
Version:        1.0
Website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Export all teams and their last activity date, Inactive days.
2. Helps to find out inactive teams based on inactive days.
3. Automatically install the Microsoft Graph PowerShell module (if not installed already) upon your confirmation.
4. The script can be executed with an MFA-enabled account too.
5. Supports Certificate-based Authentication too.
6. The script is scheduler friendly.

For detailed Script execution:   https://o365reports.com/2025/02/18/find-inactive-teams-in-microsoft-teams/
============================================================================================
#>

Param
(
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [int]$InactiveDays,
    [Switch]$ShowNeverUsedTeams,
    [Switch]$ExcludeNeverUsedTeams
    
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
  Connect-MgGraph -Scopes "Reports.Read.All"  -NoWelcome
 }
}
Connect_MgGraph


$Location=Get-Location
$TempFile="$Location\TeamsUsageReport_tempFile_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv" 
$ExportCSV = "$Location\InactiveTeams_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$Count=0
$PrintedTeams=0
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

 #Filter for showing teams with no activity
 If(($ShowNeverUsedTeams.IsPresent) -and ($LastActivityDate -ne "Never Active"))
 {
  $Print=0
 }

 #Filter for excluding never active teams
 if(($ExcludeNeverUsedTeams.IsPresent) -and ($LastActivityDate -eq "Never Active"))
 {
  $Print=0
 }



 #Export the output to CSV
 if($Print -eq 1)
 {
  $PrintedTeams++
  $ExportResult=[PSCustomObject]@{'Team Name'=$TeamName;'Team Type'=$TeamType;'Last Activity Date'=$LastActivityDate;'Inactive Days'=$InactivePeriod;'Is Deleted'=$IsDeleted;'Team Id'=$TeamId}
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
  Write-Host `nInactive teams report available in: -NoNewline -Foregroundcolor Yellow; Write-Host $ExportCSV
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
