<#
=============================================================================================
Name:           Risky Users Report in Microsoft Entra
Version:        1.0
Website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Exports all risky users in your organization to a CSV file.
2. Lists all users who have a history of risky activity.
3. Finds users based on specific risk levels and risk states.
4. Supports exporting risky users over a specified time.
5. Automatically install the Microsoft Graph PowerShell module (if not installed already) upon your confirmation.
6. The script can be executed with an MFA-enabled account too.
7. Supports Certificate-based Authentication too.
8. The script is scheduler friendly.


For detailed Script execution: https://o365reports.com/2025/05/20/how-to-get-all-risky-users-in-microsoft-entra/
============================================================================================
#>

Param
(
    [nullable[int]]$ShowRiskyUsersFromLastNDays,
    [ValidateSet("Low", "Medium", "High", "None")]
    [string[]]$RiskLevel,
    [ValidateSet("ConfirmedSafe", "Remediated", "Dismissed", "AtRisk", "ConfirmedCompromised", "None")]
    [string[]]$RiskState,
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Function Connect_MgGraph
{
 #Check for module installation
 $Module=Get-Module -Name Microsoft.Graph.Identity.SignIns -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Microsoft Graph PowerShell SDK is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Microsoft Graph PowerShell module..."
   Install-Module Microsoft.Graph.Identity.SignIns -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
  }
  else
  {
   Write-Host "Microsoft Graph PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph.Identity.SignIns cmdlet." 
   Exit
  }
 }
 #Disconnect Existing MgGraph session
 if($CreateSession.IsPresent)
 {
  Disconnect-MgGraph
 }

 Write-Host Connecting to Microsoft Graph...
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
 }
 else
 {
  Connect-MgGraph -Scopes "IdentityRiskyUser.Read.All" -NoWelcome
 }
}
Connect_MgGraph

$Location = Get-Location
$CurrentDate = Get-Date
$ExportCSV = "$Location\M365_Risky_Users_Report$($CurrentDate.ToString('yyyy-MMM-dd-ddd hh-mm-ss tt')).csv"
$Filter = @()
$ExportResult =""   
$ExportResults = @() 

if ($ShowRiskyUsersFromLastNDays -ne $null) {
    $Filter += "(RiskLastUpdatedDateTime ge $($CurrentDate.AddDays(-$ShowRiskyUsersFromLastNDays).ToString('yyyy-MM-dd')))"
} else {
    $Filter += "(RiskLastUpdatedDateTime ge $($CurrentDate.AddDays(-90).ToString('yyyy-MM-dd')))"
}

$Count=0
$PrintedLogs=0
$Filter = $Filter -join " and "
Write-Host "Generating M365 risky users' report..."
Get-MgRiskyUser -All -Filter "$($Filter)" | ForEach-Object {
 $Count++
 Write-Progress -Activity "`n     Identified $count risky users"
 $Id = $_.Id
 $RiskLastUpdatedDateTime = ($_.RiskLastUpdatedDateTime).ToLocalTime()
 $UserRiskLevel = $_.RiskLevel
 $UserRiskState = $_.RiskState
 $UserRiskDetail = $_.RiskDetail
 $UserDisplayName = $_.UserDisplayName
 $UPN = $_.UserPrincipalName
 $IsDeleted = $_.IsDeleted
 $IsProcessing = $_.IsProcessing
 $UserRiskEventType = ((Get-MgRiskyUserHistory -RiskyUserId $_.Id).Activity | Select -ExpandProperty RiskEventTypes | Select -Unique) -join ', '

 if ($UserRiskEventType -eq "") {
    $UserRiskEventType = "none"
 }

 $Print = 1

 # Apply filters based on the param values...
 if (!([string]::IsNullOrEmpty($RiskLevel)) -and ($UserRiskLevel -notin $RiskLevel)) { $Print = 0 }
 if (!([string]::IsNullOrEmpty($RiskState)) -and ($UserRiskState -notin $RiskState)) { $Print = 0 }

 #Export users to output file
 if($Print -eq 1)
 {
  $PrintedLogs++
  $ExportResult=[PSCustomObject]@{'Risk Last Updated Date Time'=$RiskLastUpdatedDateTime; 'Risky User UPN'=$UPN; 'Risky User Name'=$UserDisplayName; 'Risk Level'=$UserRiskLevel; 'Remediation Action'=$UserRiskDetail; 'Risk State'=$UserRiskState; 'Risk Event Type'=$UserRiskEventType; 'Risky User Id'=$Id; 'Is User Deleted'=$IsDeleted; 'Is Backend Processing'=$IsProcessing;}
  $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
 }
}

#Disconnect the session after execution
Disconnect-MgGraph | Out-Null

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n

#Open output file after execution
if((Test-Path -Path $ExportCSV) -eq "True")
{   
    Write-Host " The Output file availble in: " -NoNewline -ForegroundColor Yellow; Write-Host "$ExportCSV" `n 
    Write-Host " Exported report has $PrintedLogs risky user's records." 
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6)
    {
        Invoke-Item "$ExportCSV"
    }
}
else
{
    Write-Host "No logs found" 
}