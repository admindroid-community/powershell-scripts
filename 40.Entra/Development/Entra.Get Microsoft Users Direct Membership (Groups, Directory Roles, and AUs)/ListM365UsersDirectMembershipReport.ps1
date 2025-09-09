<#
=============================================================================================
Name:         List Microsoft 365 User’s Direct Membership Using PowerShell 
Version:      1.0
website:      o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script exports 3 different CSV reports.
1 - i) User's direct group membership report
1 - ii)Users with admin roles
1 - iii) Users with their administrative units
2. Retrieves guest user memberships, too. 
3. Allows you to get specific user’s direct membership within existing objects separately.   
4. You can import a CSV and filter down memberships for a list of users, too!  
5. Automatically install the required Microsoft Graph modules with your confirmation. 
6. The script can be executed with an MFA-enabled account too.  
7. Exports report results as a CSV file. 
8. The script is scheduler-friendly, making it easy to automate.
9. It supports certificate-based authentication (CBA) too. 

For detailed Script execution: : https://o365reports.com/2024/08/06/list-microsoft-365-users-direct-membership-using-powershell 
============================================================================================
#>
param(
[string]$TenantID,
[string]$ClientID,
[string]$CertificateThumbPrint,
[string]$UserId,
[string]$CSV
)

$directoryRolesFilePath = ".\Users_DirectoryRoles_Membership_Report$((Get-Date -format yyyy-MMM-dd-ddd_hh-mm-ss_tt).ToString()).csv"
$administrativeUnitsFilePath = ".\Users_AdministrativeUnits_Membership_Report$((Get-Date -format yyyy-MMM-dd-ddd_hh-mm-ss_tt).ToString()).csv"
$groupFilePath = ".\Users_GroupMembership_Report$((Get-Date -format yyyy-MMM-dd-ddd_hh-mm-ss_tt).ToString()).csv" 
$global:Count = 0
function UserDirectMembership{
  param(
   [Parameter(Mandatory = $true)]
   [PSCustomObject]$UserDetails
  )
     try {  
     $global:Count++
     Write-Progress -Activity "Getting User's Direct Membership Report.. Processed Users Count:$($Count)" -Status "Fetching User data for $($UserDetails.DisplayName)"
     Get-MgUserMemberOf -UserId $UserDetails.Id -All | Select-Object -ExpandProperty AdditionalProperties -Property id  |ForEach-Object{$DirectMembership=$_
     if ($DirectMembership.'@odata.type' -eq "#microsoft.graph.directoryRole")
                {
                    $directoryRole = [PSCustomObject]@{
                        "User Name " = $UserDetails.DisplayName
                        "UPN"=$UserDetails.UserPrincipalName
                        "DirectoryRole Name" = $DirectMembership.displayName
                        "DirectoryRole Description" = $DirectMembership.description
                        "DirectoryRole Id" = $DirectMembership.Id
                        "User SignIn Status"=$UserDetails.SignInStatus
                        "User Department"=$UserDetails.Department
                        "User Job Title"=$UserDetails.JobTitle
                  }  
         $directoryRole | Export-Csv -Path $directoryRolesFilePath -NoTypeInformation -Append -Force
            }
                elseif($DirectMembership.'@odata.type' -eq "#microsoft.graph.group")
                {

                    $group = [PSCustomObject]@{
                        "User Name " = $UserDetails.DisplayName
                        "UPN"=$UserDetails.UserPrincipalName
                        "Group Name" = $DirectMembership.displayName
                        "Group Description" = $DirectMembership.description
                        "Group Visibility" = $DirectMembership.visibility
                        "Group Mail Id" = $DirectMembership.mail
                        "Group Types" = ""
                        "Group Created Date Time" = Get-Date -Date $DirectMembership.createdDateTime
                        "Group Id" = $DirectMembership.Id
                        "User SignIn Status"=$UserDetails.SignInStatus
                        "User Department"=$UserDetails.Department
                        "User Job Title"=$UserDetails.JobTitle
                       

                    }

                    if ($DirectMembership.groupTypes[0] -eq "Unified")
                    {
                        $group."Group Types" = "Microsoft 365 group"
                    }
                    elseif($DirectMembership.securityEnabled -and $DirectMembership.mailEnabled)
                    {
                        $group."Group Types" = "Mail-enabled security group"
                    }
                    elseif($DirectMembership.securityEnabled)
                    {
                        $group."Group Types" = "Security group"
                    }
                    else
                    {
                        $group."Group Types" = "Distribution list"
                    }
                     
                     $group| Export-Csv -Path $groupFilePath -NoTypeInformation -Append -Force
                }
                elseif($DirectMembership.'@odata.type' -eq "#microsoft.graph.administrativeUnit")
                {
                    $administrativeUnits = [PSCustomObject]@{
                        "User Name " = $UserDetails.DisplayName
                        "UPN"=$UserDetails.UserPrincipalName
                        "AU Name" = $DirectMembership.displayName
                        "AU Description" = $DirectMembership.description
                        "AU Id" = $DirectMembership.Id
                        "User SignIn Status"=$UserDetails.SignInStatus
                        "User Department"=$UserDetails.Department
                        "User Job Title"=$UserDetails.JobTitle
                    }
                  
                  $administrativeUnits | Export-Csv -Path $administrativeUnitsFilePath -NoTypeInformation -Append -Force
                }

                                                                                                                                                
           }
}
 catch {
        Write-Host "Error occurred: $( $_.Exception.Message )" -ForegroundColor Red
        Exit
    }
}



#Module installation 
$Module = Get-Module -Name Microsoft.Graph.Users -ListAvailable
if ($Module.count -eq 0)
{
    Write-Host Microsoft.Graph.Users is not available in Your System -ForegroundColor Red
    $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No
    if ($Confirm -eq "y" -or $Confirm -eq "Y")
    {
        try
        {
            Install-Module Microsoft.Graph.Users -Force -AllowClobber -Scope CurrentUser
        }
        catch
        {
            Write-Host "Error occurred : $( $_.Exception.Message )" -ForegroundColor Red
            Exit
        }
        Write-Host Microsoft.Graph.Users installed successfully...  -ForegroundColor Green
      
    }
    else
    {
        Write-Host Microsoft.Graph.Users is required .Please Install-Module Microsoft.Graph.Users to continue..
        Exit
    }
}
#Authenication
try
{
       
if (($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbPrint -ne "")) {
$Connect = Connect-MgGraph -TenantId $TenantID.Trim() -ClientID $ClientID.Trim() -CertificateThumbprint $CertificateThumbPrint.Trim() -ErrorAction Stop
}else{    
     $Connect = Connect-MgGraph -Scopes "Directory.Read.All" -ErrorAction Stop}
}
catch
{
    Write-Host "Error occurred while connecting to Microsoft Graph: $( $_.Exception.Message )" -ForegroundColor Red
    Exit
}
function UserClass{
   param(
   [Parameter(Mandatory = $true)]
   [string]$Userid
  )
  try{
     $User=Get-MgUser -UserId $Userid.Trim() -Property DisplayName,UserPrincipalName,AccountEnabled,department,JobTitle,id | Select-Object  DisplayName,UserPrincipalName,AccountEnabled,department,JobTitle,id
     $UserDetails=[PSCustomObject]@{
     "DisplayName"=$User.DisplayName
     "UserPrincipalName"=$User.UserPrincipalName
     "Department"=$User.Department
     "JobTitle"=$User.JobTitle
     "Id"=$User.Id
     "SignInStatus"=""
     }
     if($User.AccountEnabled){
      $UserDetails.SignInStatus="Enabled"}
      else{
      $UserDetails.SignInStatus="Disabled"}
      UserDirectMembership -UserDetails $UserDetails
  }catch{
   Write-Host "Error occurred : $( $_.Exception.Message )" -ForegroundColor Red
  }
}


#Get membership details for a single user
if($UserId -ne "")
   {
   UserClass -Userid $UserId
   } 

#Get membership details for a list of users
   elseif($CSV -ne "")
   {
      if ((Test-Path -Path $CSV) -eq "True") {
                
                Import-Csv -Path $CSV | ForEach-Object {
                   $UserId = $_.UserId 
                   UserClass -Userid $UserId
                  }
               }
                else {
                Write-Host "Incorrect Csv File Path : $CSV" -ForegroundColor Red
                Exit
            } 
   }

#Get membership details for all users
else{
    try{
 
   Get-MgUser -All -Property DisplayName,UserPrincipalName,AccountEnabled,department,JobTitle,id | Select-Object  DisplayName,UserPrincipalName,AccountEnabled,department,JobTitle,id |ForEach-Object{$User=$_
     $UserDetails=[PSCustomObject]@{
     "DisplayName"=$User.DisplayName
     "UserPrincipalName"=$User.UserPrincipalName
     "Department"=$User.Department
     "JobTitle"=$User.JobTitle
     "Id"=$User.Id
     "SignInStatus"=""
     }
     if($User.AccountEnabled){
      $UserDetails.SignInStatus="Enabled"}
      else{
      $UserDetails.SignInStatus="Disabled"}
   
      UserDirectMembership -UserDetails $UserDetails}
  }catch{
   Write-Host "Error occurred : $( $_.Exception.Message )" -ForegroundColor Red
  }
}

 Write-Host `n Script executed successfully -ForegroundColor Green
 if ((Test-Path -Path $groupFilePath) -eq "True") {
     Write-Host `n "Users' group membership report availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$groupFilePath" `n 
 }else{ $groupFilePath=""
 Write-Host `n "No data is available for the Users' group membership report" -NoNewline `n} 
 if ((Test-Path -Path $directoryRolesFilePath) -eq "True") {
     Write-Host `n "Users' directory role membership report availble in :" -NoNewline -ForegroundColor Yellow; Write-Host "$directoryRolesFilePath" `n 
 }else{$directoryRolesFilePath=""
  Write-Host `n "No data is available for the Users directory role membership report" -NoNewline `n
 }
 
  if ((Test-Path -Path $administrativeUnitsFilePath) -eq "True") {
      Write-Host `n "Users' Administrative Units membership report availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$administrativeUnitsFilePath" `n 
  }else{ $administrativeUnitsFilePath=""
     Write-Host `n "No data is available for the Users' Administrative Units membership report" -NoNewline `n
  }
  
      
       if($directoryRolesFilePath -ne "" -or $administrativeUnitsFilePath -ne "" -or  $groupFilePath -ne ""){
        $Prompt = New-Object -ComObject wscript.shell  
        $UserInput = $Prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)  
        if ($UserInput -eq 6) {  
        if($groupFilePath -ne ""){
           Invoke-Item  $groupFilePath
           }
           if($directoryRolesFilePath -ne ""){
           Invoke-Item $directoryRolesFilePath
           }
           if($administrativeUnitsFilePath -ne ""){
           Invoke-Item $administrativeUnitsFilePath
           }
            


        }}
Disconnect-MgGraph |Out-Null
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n

