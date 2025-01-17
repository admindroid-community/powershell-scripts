<#
=============================================================================================
Name:           Find Teams with Guests and External Users in Microsoft Teams     
Version:        1.0
Website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~
1. The script exports 2 different reports: 
     -Detailed report: Teams and their guest user details  
     -Summary report: Teams and their guest users count  
2. Automatically installs the Microsoft Teams PowerShell module (if not installed already) upon your confirmation. 
3. Exports the report result to CSV. 
4. The script can be executed with an MFA enabled account too. 
5. It can be executed with certificate-based authentication (CBA) too. 
6. The script is schedular-friendly.

For detailed Script execution: https://o365reports.com/2025/01/14/find-teams-with-guests-and-external-users-in-microsoft-teams/
============================================================================================
#>

#params 
param(
[string]$UserName, 
[string]$Password, 
[string]$TenantId,
[string]$AppId,
[string]$CertificateThumbprint
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
  Import-Module MicrosoftTeams
 }
 else
 {
  Write-Host MicrosoftTeams module is required.Please install module using Install-Module MicrosoftTeams cmdlet.
  Exit
 }
}


#Connect to MS Teams
 #Storing credential in script for scheduling purpose/ Passing credential as parameter
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  $Team=Connect-MicrosoftTeams -Credential $Credential
 }
 elseif(($TenantId -ne "") -and ($AppId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  $Team=Connect-MicrosoftTeams  -TenantId $TenantId -ApplicationId $AppId -CertificateThumbprint $CertificateThumbprint 
 }
 else
 {  
  $Team=Connect-MicrosoftTeams
 }

#Check for Teams connectivity
If($Team -ne $null)
{
 Write-host `nSuccessfully connected to Microsoft Teams. 
}
else
{
 Write-Host Error occurred while creating Teams session. Please try again -ForegroundColor Red
 exit
}

$Result=""  
$Results=@() 
$Count=0
$PrintedTeams=0
$Location=Get-Location
$ExportCSV = "$Location\FindTeams_with_Guests$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult="$Location\Get_All_Guests_in_Teams$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"

Write-Host Exporting all teams and their guests...

Get-Team | foreach {
 $TeamName=$_.DisplayName
 $PrintTeam=0
 $GuestCount=0
 $Count++
 
 Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName"
 $Count++
 $GroupId=$_.GroupId
 Get-TeamUser -GroupId $GroupId -Role Guest | foreach {
  $PrintTeam=1
  $GuestCount++
  $Name=$_.Name
  $MemberMail=$_.User
  $Result=[PSCustomObject]@{'Teams Name'=$TeamName;'Guest Name'=$Name;'Guest Mail'=$MemberMail}
  $Result | Export-Csv $ExportResult -NoTypeInformation -Append
 }
 if($PrintTeam -eq "1")
 {
  $PrintedTeams++
  $ExportResults=[PSCustomObject]@{'Team Name'=$TeamName;'Guest Count'=$GuestCount}
  $ExportResults | Export-Csv -Path $ExportCSV -Notype -Append
 }
}
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
Write-Host $PrintedTeams out of $Count teams contain guests

#Open output file after execution 
if((Test-Path -Path $ExportCSV) -eq "True") 
{
 Write-Host `nThe Output files available in:  -NoNewline -ForegroundColor Cyan
 Write-Host $Location 
 $Prompt = New-Object -ComObject wscript.shell      
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
 If ($UserInput -eq 6)   
 {   
  Invoke-Item "$ExportCSV"   
  Invoke-Item $ExportResult
 } 
}
