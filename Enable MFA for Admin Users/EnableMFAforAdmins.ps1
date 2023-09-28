<#
=============================================================================================
Name:           Enable MFA for all Office 365 admins
Version:        1.0
Website:        m365scripts.com

Script Highlights:
~~~~~~~~~~~~~~~~~
1.Finds admins without MFA and enables MFA for them.
2.Allows to enable MFA for licensed admins alone.
3.Exports MFA enabling status to CSV file.
4.The script can be executed with MFA enabled account.
5.Credentials are passed as parameters, so worry not!

For detailed script execution: https://m365scripts.com/security/enabling-mfa-for-admins-using-powershell/
============================================================================================
#>




#PARAMETERS
param ( 
[String] $UserName = $null, 
[String] $Password = $null,
[Switch] $LicensedAdminsOnly
)

#Check for Module Availability
$MsOnline = (Get-Module MsOnline -ListAvailable).Name 
if($MsOnline -eq $null)
{ 
   Write-Host "Important: Module MsOnline is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
   $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
   if($Confirm -match "[yY]")
     { 
       Write-Host "Installing MsOnline module..."
       Install-Module MsOnline -Repository PsGallery -Force -AllowClobber 
       Write-Host "Required Module is installed in the machine Successfully" -ForegroundColor Magenta 
     } 
    else
     { 
       Write-Host "Exiting. `nNote: MsOnline module must be available in your system to run the script" 
       Exit 
     } 
}


#Importing Module by default will avoid the cmdlet unrecognized error 
Import-Module MsOnline -Force 

#CONNECTING TO MSOLSERVICE.......
Write-Host "Connecting to Msolservice..."`n
if(($UserName -ne "") -and ($Password -ne ""))
{
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  Connect-MsolService -Credential $Credential 
}
else
{
  Connect-MsolService
}

#Creating Object for Enable MFA
$MultiFactorAuthentication_Object= New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$MultiFactorAuthentication_Object.RelyingParty = "*"
$MultiFactorAuthentication_Object.State = "Enabled"
$MultiFactorAuthentication = @($MultiFactorAuthentication_Object)


#Separating Admin without MFA And Enable MFA for them
Write-Host "Preparing Admin Without MFA List And Enable MFA for them..."`n
$OutputCsv=".\AdminsWithoutMFAReport_$((Get-Date -format MMM-dd` hh-mm` tt).ToString()).csv"
$global:CountForSuccess = 0
$global:CountForFailed = 0


#function for enable MFA for Admins
function EnableMFAforadmin
{
  $AdminName = $User.DisplayName
  $LicensedStatus = if($User.isLicensed) { "Licensed" } else { "UnLicensed" }
  
  try
  {
    Set-MsolUser -UserPrincipalName $User.userprincipalname -StrongAuthenticationRequirements $MultiFactorAuthentication -ErrorAction Stop
    $global:CountForSuccess++
    $MFAstatus = "MFA successfully Assigned"
  }
  catch 
  {
    $global:CountForFailed++
    $MFAstatus = "Failed To Assign MFA"
  }
  $User = @{'Admin Name'=$AdminName;'UPN' =$User.UserPrincipalName;'Roles'=($Roles.Name)-join',';'License Status'=$LicensedStatus;'MFA Status'=$MFAstatus}
  $ExportUser = New-Object PSObject -Property $User
  $ExportUser | Select-Object 'Admin Name','UPN','Roles','License Status','MFA Status' | Export-csv -path $OutputCsv -NoType -Append
  Write-Progress -Activity "Updating $Adminname ..." -Status "MFA Successfully Assigned for $CountForSuccess Admins , Failed for $CountForFailed Admins"
}



#Filter Admin User Using MsolUserRole
Get-MsolUser -All | Select UserPrincipalName,DisplayName,StrongAuthenticationRequirements,isLicensed | ForEach-Object {

 $User = $_
 $Roles = (Get-MsolUserRole -UserPrincipalName $User.UserPrincipalName) 
 if($LicensedAdminsOnly.IsPresent)
  {
    if($Roles.Name -ne $null -and $User.StrongAuthenticationRequirements.State -eq $null -and $User.IsLicensed -eq $true)
    {
       EnableMFAforadmin
    }
  }
 else
  {
    if($Roles.name -ne $null -and $User.StrongAuthenticationRequirements.State -eq $null)
    {
       EnableMFAforadmin
    }
  }
}


#Display Details about succesfull and failure 
if($CountForSuccess -ne 0 -or $CountForFailed -ne 0)
 {
   Write-Host "MFA Successfully Enabled for $CountForSuccess Admins and MFA Failed for $CountForFailed Admins"
 }
 else
 { 
   Write-Host "Already All the Admins are enabled MFA"`n`n
 } 


#Open output file after execution 
if((Test-Path -Path $OutputCsv) -eq "True") { 
 Write-Host `n "The Output file availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$outputCsv"
 $Prompt = New-Object -ComObject wscript.shell    
 $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)    
 If ($UserInput -eq 6)    
  {    
   Invoke-Item "$OutputCSV"    
  }
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
} 