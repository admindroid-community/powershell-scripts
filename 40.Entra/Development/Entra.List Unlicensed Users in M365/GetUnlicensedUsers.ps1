<#
=============================================================================================
Name:           Find Unlicensed Users in Microsoft 365 using PowerShell 
Version:        1.0
website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~

1. The script uses MS Graph PowerShell and installs MS Graph PowerShell SDK (if not installed already) upon your confirmation. 
2. The script can be executed with MFA enabled account too. 
3. Exports both disabled and enabled user accounts without licenses. 
4. Exports unlicensed member accounts only, excluding guests. 
5. Identifies unlicensed users within specific departments. 
6. Filters unlicensed users based on job title. 
7. Exports report results as a CSV file. 
8. The script is scheduler friendly. 
9. It can be executed with certificate-based authentication (CBA) too. 

For detailed Script execution:  https://o365reports.com/2024/08/20/find-unlicensed-users-in-microsoft-365-using-powershell/
============================================================================================
#>
Param
(
    [switch]$IncludeDisabledUsers,
    [Switch]$ExcludeGuests,
    [switch]$CreateSession,
    [string]$Department,
    [string]$JobTitle,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Function Connect_MgGraph
{
 #Check for module installation
 $Module=Get-Module -Name Microsoft.Graph.Beta -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Microsoft Graph PowerShell SDK is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Microsoft Graph PowerShell module..."
   Install-Module Microsoft.Graph.beta -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
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
  Disconnect-MgGraph
 }

 #Connecting to MgGraph beta
 Write-Host Connecting to Microsoft Graph...
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
 }
 else
 {
  Connect-MgGraph -Scopes "User.Read.All"  -NoWelcome
 }
}
Connect_MgGraph

$Location=Get-Location
$ExportCSV="$Location\UnlicensedUsers_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult=""   
$ExportResults=@()  


$Count=0
$PrintedUser=0
#retrieve users
$RequiredProperties=@('UserPrincipalName','CreatedDateTime','AccountEnabled','Department','JobTitle','UserType')
Get-MgBetaUser -Filter 'assignedLicenses/$count eq 0' -ConsistencyLevel eventual -CountVariable unlicensedUserCount -All -Property $RequiredProperties | select $RequiredProperties | ForEach-Object {
 $Count++
 $UPN=$_.UserPrincipalName
 Write-Progress -Activity "`n     Processed user: $Count - $UPN"
 $CreatedDate=$_.CreatedDateTime
 $AccountEnabled=$_.AccountEnabled
 $Dept=$_.Department
 $Title=$_.JobTitle
 $UserType=$_.UserType

 if($AccountEnabled -eq $true)
 {
  $AccountStatus='Enabled'
 }
 else
 {
  $AccountStatus='Disabled'
 }

 

 #Inactive days based on interactive signins filter
 if(!($IncludeDisabledUsers.IsPresent) -and ($AccountStatus -eq 'Disabled'))
 {
  return
 }

 if(($ExcludeGuests.IsPresent) -and ($UserType -eq 'Guest'))
 {
  return
 }
    
 if(($Department -ne "") -and ($Department -ne $Dept))
 {
  return
 }

 If(($JobTitle -ne "") -and ($Title -ne $JobTitle))
 {
  return
 }

 #Export users to output file
  
 $PrintedUser++
 $ExportResult=[PSCustomObject]@{'UPN'=$UPN;'Department'=$Dept;'Job Title'=$Title;'Creation Time'=$CreatedDate;'User Type'=$UserType;'Account Status'=$AccountStatus;}
 $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
 
}

#Open output file after execution
Write-Host `nScript executed successfully
if((Test-Path -Path $ExportCSV) -eq "True")
{
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    Write-Host "Exported report has $PrintedUser user(s)" 
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6)
    {
        Invoke-Item "$ExportCSV"
    }
    Write-Host "Exported report available in: $ExportCSV"
}
else
{
    Write-Host "No user found" -ForegroundColor Red
}