<#
=============================================================================================
Name:           Export Office 365 User's Manager Report
Description:    This script exports Office 365 users and their manager to CSV
Version:        2.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.Generates 10+ different manager reports to view the managers and direct reports status.  
2.Automatically installs MS Graph module upon your confirmation when it is not available in your machine. 
3.Shows list of all Azure AD users and their manager.  
4.List of all Office 365 users with no manager. 
5.Allows specifying user departments to get their manager details. 
6.You can get the direct reports of the Office 365 managers. 
7.Supports both MFA and Non-MFA accounts.    
8.Exports the report in CSV format.  
9.Scheduler-friendly. 
10.Supports certificate-based authentication (CBA) too. 

For detailed Script execution: https://o365reports.com/2021/07/13/export-office-365-user-manager-and-direct-reports-using-powershell
============================================================================================
#>

param (
    [string] $UserName = $null,
    [string] $Password = $null,
    [Switch] $UsersWithoutManager,
    [Switch] $DisabledUsers,
    [Switch] $UnlicensedUsers,
    [Switch] $DirectReports,
    [string[]] $Department,
     [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)


Function Connect_MgGraph
{
 #Check for module installation
 $Module=Get-Module -Name microsoft.graph.beta -ListAvailable
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
   Write-Host "Microsoft Graph Beta PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph cmdlet." 
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
  Connect-MgGraph -Scopes "User.Read.All","AuditLog.read.All"  -NoWelcome
 }
}

#Handle Empty attributes here
Function GetPrintableValue($RawData) {
    if (($null -eq $RawData) -or ($RawData.Equals(""))) {
        return "-";
    } else {
        $StringVal = $RawData | Out-String
        return $StringVal;
    }
}

#Processes the param and prepares respective filters.
Function FindUseCase {
    #Appends all the usecases choice of the user
    if (($Department.Length) -gt 0) {
        $DepartmentList = '"' + ($Department -join '","') + '"'
        $UseCaseFilter = '$_.Department -in ' + $DepartmentList
    }
    if ($DisabledUsers.IsPresent) {
        if ($UseCaseFilter -ne $null) {
            $UseCaseFilter = $UseCaseFilter.ToString() + '-and $_.AccountEnabled -eq $false'
        }
        else {
            $UseCaseFilter = '$_.AccountEnabled -eq $false'
        }
    }
    if ($UnlicensedUsers.IsPresent) {
        if ($UseCaseFilter -ne $null) {
            $UseCaseFilter = $UseCaseFilter.ToString() + ' -and ($_.AssignedLicenses).count -eq 0'
        } else {
            $UseCaseFilter = '($_.AssignedLicenses).count -eq 0'
        }
    }
    Write-Host "Generating report..." -ForegroundColor Cyan
    if ($UseCaseFilter -ne $null) { 
        #Filters the users to generate report
        $UseCaseFilter = [ScriptBlock]::Create($UseCaseFilter)
        Get-MgBetaUser -All | Where-Object $UseCaseFilter | foreach-object {
            $CurrUserData = $_
            $UserId=$_.Id
            ProcessUserData
        }
    } else { 
        #No Filter- Gets all the users without any filter
        Get-MgBetaUser -All | foreach-object {
            $CurrUserData = $_
            $UserId=$_.Id
            ProcessUserData
        }
    }
}

#Processes User info and calls respective functions based on the requested report
Function ProcessUserData {
    if ($DirectReports.IsPresent) {
        $CurrUserDirectReport = Get-MgBetaUserDirectReport -UserId $userId | select -ExpandProperty additionalProperties
        if ($CurrUserDirectReport -ne $Empty) { 
            #Manager has Direct Reports, Exporting Manager and their Direct Reports Info
            RetrieveUserDirectReport
            ExportManagerAndDirectReports
        }
    } else {
        $CurrManagerData = Get-MgBetaUserManager -UserId $userId -ErrorAction SilentlyContinue
        #Processing Manager Report Types
        if ($CurrManagerData -ne $Empty -and !$UsersWithoutManager.IsPresent) {
            #User has Manager assigned, Exporting User & Manager Data.
            RetrieveUserManagerData
            ExportUserAndManagerData
        }
        if ($CurrManagerData -eq $Empty -and $UsersWithoutManager.IsPresent) {
            #User has no Manager, Exporting User Data only
            RetrieveUserManagerData
            ExportUserDataOnly  
        }
    }
}

#Saves User and Manager info into variables
Function RetrieveUserManagerData {
    #Processing user data
    $global:ExportedUser = $global:ExportedUser + 1
    $global:UserName = $CurrUserData.DisplayName
    $global:UserUPN = $CurrUserData.UserPrincipalName
    $global:UserAccountType = $CurrUserData.UserType
    $global:UserDepartment = GetPrintableValue $CurrUserData.Department

    if (($CurrUserData.AssignedLicenses) -ne $null) {
        $global:UserLicense = "Licensed"
    }
    else {
        $global:UserLicense = "Unlicensed"

    }
    if ( ($CurrUserData.AccountEnabled) -eq $True) {
        $global:UserAccount = "Active"
    }
    else {
        $global:UserAccount = "Disabled"
    }

    #Processing manager data 
    if ($CurrManagerData -ne $Empty) {
        $ManagerDetails=$CurrManagerData.AdditionalProperties
        $global:ManagerName = $ManagerDetails.displayName
        $global:ManagerUPN = $ManagerDetails.userPrincipalName
        $global:ManagerDepartment = GetPrintableValue $ManagerDetails.department
        if ( ($ManagerDetails.accountEnabled) -eq $True) {
            $global:ManagerAccount = "Active"
        }
        else {
            $global:ManagerAccount = "Disabled"
        }
    }
}

#Saves Manager and Direct Reports info into variables
Function RetrieveUserDirectReport {
    #Pocessing manager data
    $global:ExportedUser = $global:ExportedUser + 1
    $global:ManagerName = $CurrUserData.DisplayName
    $global:ManagerUPN = $CurrUserData.UserPrincipalName
    $global:ManagerDepartment = GetPrintableValue $CurrUserData.Department

    #Processing Direct report data
    $global:NoOfDirectReports = ($CurrUserDirectReport.displayName).count
    Write-Host $CurrUserDirectReport -ForegroundColor Green
    $global:DirectReportsNames=$CurrUserDirectReport.displayName -join ","
    $global:DirectReportsUPNs=$CurrUserDirectReport.userPrincipalName -join ","

}

#Used for 'UsersWithoutManager' param. Exports user info alone.
Function ExportUserDataOnly {
    $global:ExportCSVFileName = "UsersWithoutManagerReport-" + $global:ReportTime 
    Write-Progress "Retrieving the Data of the User: $global:UserName" "Processed Users Count: $global:ExportedUser"

    $ExportResult = @{'User Name' = $global:UserName; 'UPN' = $global:UserUPN; 'Account Status' = $global:UserAccount; 'User Type' = $global:UserAccountType; 'License Status' = $global:UserLicense; 'Department' = $global:UserDepartment }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'User Name', 'UPN', 'Department', 'User Type', 'Account Status', 'License Status' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force 
}

#Used for 'RetrieveUserManagerData' param. Exports User and Manager info.
Function ExportUserAndManagerData {
    $global:ExportCSVFileName = "UsersWithManagerReport-" + $global:ReportTime 
    Write-Progress "Retrieving the Manager Data of the User: $global:UserName" "Processed Users Count: $global:ExportedUser"

    $ExportResult = @{'User Name' = $global:UserName; 'User UPN' = $global:UserUPN; 'User Account Status' = $global:UserAccount; 'User Account Type' = $global:UserAccountType; 'Manager Name' = $global:ManagerName; 'Manager UPN' = $global:ManagerUPN ; 'Manager Department' = $global:ManagerDepartment; 'Manager Account Status' = $global:ManagerAccount ; 'User Department' = $global:UserDepartment; 'User License Status' = $global:UserLicense }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'User Name', 'User UPN', 'Manager Name', 'Manager UPN', 'Manager Department', 'Manager Account Status', 'User Department', 'User Account Status', 'User Account Type', 'User License Status' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force 
}

#Used for 'RetrieveUserManagerData' param. Exports User and Direct reports.
Function ExportManagerAndDirectReports {
    $global:ExportCSVFileName = "UsersWithDirectReports-" + $global:ReportTime 
    Write-Progress "Retrieving the Manager Data of: $global:ManagerName" "Processed Managers Count: $global:ExportedUser"

    $ExportResult = @{'Manager Name' = $global:ManagerName; 'Manager UPN' = $global:ManagerUPN; 'Manager Department' = $global:ManagerDepartment; 'No. of Direct Reports' = $global:NoOfDirectReports; 'Direct Reports Names' = $global:DirectReportsNames; 'Direct Reports UPN' = $global:DirectReportsUPNs}
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Manager Name' , 'Manager UPN' , 'Manager Department' , 'No. of Direct Reports' , 'Direct Reports Names' , 'Direct Reports UPN' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
}

# Execution starts here.
Connect_MgGraph

$global:ExportedUser = 0
$global:ReportTime = ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"

FindUseCase

if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") {     
    #Open file after code execution finishes
    Write-Host "The output file available in:" -NoNewline -ForegroundColor Yellow
	Write-Host .\$global:ExportCSVFileName `n
    write-host "Exported $global:ExportedUser records to CSV." 
    $prompt = New-Object -ComObject wscript.shell    
    $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
    If ($userInput -eq 6) {    
        Invoke-Item "$global:ExportCSVFileName"
    }  
} else {
    #Notification when usecase doesn't have the data in the tenant
    Write-Host `n"No data found with the specified criteria"
}
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
