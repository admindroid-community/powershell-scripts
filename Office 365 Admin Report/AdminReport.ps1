<#
=============================================================================================
Name:           Microsoft 365 Admin Report
Description:    This script exports Microsoft 365 admin role group membership to CSV
Version:        1.0
website:        o365reports.com
Script by:      O365Reports Team
For detailed Script execution: https://o365reports.com/2021/03/02/Export-Office-365-admin-role-report-powershell
============================================================================================
#>

param ( 
[string] $UserName = $null, 
[string] $Password = $null, 
[switch] $RoleBasedAdminReport, 
[String] $AdminName = $null, 
[String] $RoleName = $null) 

#Check for module availability
$msOnline = (get-module MsOnline -ListAvailable).Name 
if($msOnline -eq $null){ 
Write-host "Important: Module MsOnline is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
$confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No  
if($confirm -match "[yY]") { 
Write-host "Installing MsOnline module..."
Install-Module MsOnline -Repository PsGallery -Force -AllowClobber 
Write-host "Required Module is installed in the machine Successfully" -ForegroundColor Magenta 
 } elseif($confirm -cnotmatch "[yY]" ){ 
Write-host "Exiting. `nNote: MsOnline module must be available in your system to run the script" 
Exit 
  } 
}

#Importing Module by default will avoid the cmdlet unrecognized error 
Import-Module MsOnline -Force 
Write-Host "Connecting to Office 365..." 

#Storing credential in script for scheduling purpose/Passing credential as parameter   
if(($UserName -ne "") -and ($Password -ne ""))   
{   
$securedPassword = ConvertTo-SecureString -AsPlainText $Password -Force   
$credential  = New-Object System.Management.Automation.PSCredential $UserName,$securedPassword   
Connect-MsolService -Credential $credential | Out-Null 
}   
else   
 {   
Connect-MsolService 
 }  

Write-Host "Preparing admin report..." 
$admins=@() 
$list = @() 
$outputCsv=".\AdminReport_$((Get-Date -format MMM-dd` hh-mm` tt).ToString()).csv" 

function process_Admin{ 
$roleList= (Get-MsolUserRole -UserPrincipalName $admins.UserPrincipalName | Select-Object -ExpandProperty Name) -join ',' 
if($admins.IsLicensed -eq $true)
 { 
$licenseStatus = "Licensed" 
 }
else
  { 
$licenseStatus= "Unlicensed" 
  } 
if($admins.BlockCredential -eq $true)
 { 
$signInStatus = "Blocked" 
 }
else
  { 
$signInStatus = "Allowed" 
  } 
$displayName= $admins.DisplayName 
$UPN= $admins.UserPrincipalName 
Write-Progress -Activity "Currently processing: $displayName" -Status "Updating CSV file"
if($roleList -ne "") 
 { 
$exportResult=@{'AdminEmailAddress'=$UPN;'AdminName'=$displayName;'RoleName'=$roleList;'LicenseStatus'=$licenseStatus;'SignInStatus'=$signInStatus} 
$exportResults= New-Object PSObject -Property $exportResult         
$exportResults | Select-Object 'AdminName','AdminEmailAddress','RoleName','LicenseStatus','SignInStatus' | Export-csv -path $outputCsv -NoType -Append  
  } 
} 

function process_Role{ 
$adminList = Get-MsolRoleMember -RoleObjectId $roles.ObjectId #Email,DisplayName,Usertype,islicensed 
$displayName = ($adminList | Select-Object -ExpandProperty DisplayName) -join ',' 
$UPN = ($adminList | Select-Object -ExpandProperty EmailAddress) -join ',' 
$RoleName= $roles.Name 
Write-Progress -Activity "Processing $RoleName role" -Status "Updating CSV file"
if($displayName -ne "")
 { 
$exportResult=@{'RoleName'=$RoleName;'AdminEmailAddress'=$UPN;'AdminName'=$displayName} 
$exportResults= New-Object PSObject -Property $exportResult 
$exportResults | Select-Object 'RoleName','AdminName','AdminEmailAddress' | Export-csv -path $outputCsv -NoType -Append 
 } 
} 

#Check to generate role based admin report
if($RoleBasedAdminReport.IsPresent)
{ 
Get-MsolRole | ForEach-Object { 
$roles= $_        #$ObjId = $_.ObjectId;$_.Name 
process_Role 
 } 
}

#Check to get admin roles for specific user
elseif($AdminName -ne "")
{ 
$allUPNs = $AdminName.Split(",") 
ForEach($admin in $allUPNs) 
 { 
$admins = Get-MsolUser -UserPrincipalName $admin -ErrorAction SilentlyContinue 
if( -not $?)
  { 
Write-host "$admin is not available. Please check the input" -ForegroundColor Red 
  }
else
  { 
process_Admin 
  } 
 } 
}

#Check to get all admins for a specific role
elseif($RoleName -ne "")
{ 
$RoleNames = $RoleName.Split(",") 
ForEach($name in $RoleNames) 
 { 
$roles= Get-MsolRole -RoleName $name -ErrorAction SilentlyContinue 
if( -not $?)
  { 
Write-Host "$name role is not available. Please check the input" -ForegroundColor Red 
  }
else
  { 
process_Role 
  } 
 } 
}

#Generating all admins report
else
 { 
Get-MsolUser -All | ForEach-Object  { 
$admins= $_ 
process_Admin 
 } 
} 
write-Host "`nThe script executed successfully" 

#Open output file after execution 
if((Test-Path -Path $outputCsv) -eq "True") { 
Write-Host "The Output file availble in $outputCsv" -ForegroundColor Green 
$prompt = New-Object -ComObject wscript.shell    
$userInput = $prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)    
If ($userInput -eq 6)    
 {    
Invoke-Item "$OutputCSV"    
 }  
} 
                                                