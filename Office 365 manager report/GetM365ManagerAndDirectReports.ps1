
param (
    [string] $UserName = $null,
    [string] $Password = $null,
    [Switch] $UsersWithoutManager,
    [Switch] $DisabledUsers,
    [Switch] $UnlicensedUsers,
    [Switch] $DirectReports,
    [string[]] $Department
)


#Check AzureAD module availability and connects the module
Function ConnectToAzureAD {
    $AzureAd = (Get-Module AzureAD -ListAvailable).Name
    if ($Empty -eq $AzureAd) {
        Write-host "Important: AzureAD PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."  
        $confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
        if ($confirm -match "[yY]") { 
            Write-host "Installing AzureAD"
            Install-Module AzureAd -Allowclobber -Repository PSGallery -Force
            Write-host "AzureAD module is installed in the system successfully."
        }
        else { 
            Write-host "Exiting. `nNote: AzureAD PowerShell module must be available in your system to run the script."  
            Exit 
        }
    }
    #Importing Module by default will avoid the cmdlet unrecognized error 
    Import-Module AzureAd -ErrorAction SilentlyContinue -Force
    #Storing credential in script for scheduling purpose/Passing credential as parameter   
    if (($UserName -ne "") -and ($Password -ne "")) {   
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force   
        $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword   
        Connect-AzureAD -Credential $Credential | Out-Null
    }   
    else {   
        Connect-AzureAD | Out-Null
    }
    Write-Host "AzureAD PowerShell module is connected successfully"
    #End of Connecting AzureAD
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

    if ($UseCaseFilter -ne $null) { 
        #Filters the users to generate report
        $UseCaseFilter = [ScriptBlock]::Create($UseCaseFilter)
        Get-AzureADUser | Where-Object $UseCaseFilter | foreach-object {
            $CurrUserData = $_
            ProcessUserData
        }
    } else { 
        #No Filter- Gets all the users without any filter
        Get-AzureADUser | foreach-object {
            $CurrUserData = $_
            ProcessUserData
        }
    }
}

#Processes User info and calls respective functions based on the requested report
Function ProcessUserData {
    if ($DirectReports.IsPresent) {
        $CurrUserDirectReport = Get-AzureADUserDirectReport -ObjectID (($CurrUserData.ObjectId).Tostring())
        if ($CurrUserDirectReport -ne $Empty) { 
            #Manager has Direct Reports, Exporting Manager and their Direct Reports Info
            RetrieveUserDirectReport
            ExportManagerAndDirectReports
        }
    } else {
        $CurrManagerData = Get-AzureADUserManager -objectID (($CurrUserData.ObjectId).Tostring())

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
        $global:ManagerName = $CurrManagerData.DisplayName
        $global:ManagerUPN = $CurrManagerData.UserPrincipalName
        $global:ManagerDepartment = GetPrintableValue $CurrManagerData.Department
        if ( ($CurrManagerData.AccountEnabled) -eq $True) {
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
    $global:NoOfDirectReports = ($CurrUserDirectReport.DisplayName).count
    if ($global:NoOfDirectReports -gt 1) {
        $NameList = @()
        $UPNList = @()

        $CurrUserDirectReport | Select-Object DisplayName | ForEach-Object { 
            $NameList += $($_.DisplayName)
            $global:DirectReportsNames = ($NameList -join ", ") 
        }
        $CurrUserDirectReport | Select-Object UserPrincipalName | ForEach-Object { 
            $UPNList += $($_.UserPrincipalName)
            $global:DirectReportsUPNs = ($UPNList -join ", ") 
        }
        
    } elseif ($global:NoOfDirectReports.count -eq 1) {
        $global:DirectReportsNames = $CurrUserDirectReport.DisplayName
        $global:DirectReportsUPNs = $CurrUserDirectReport.UserPrincipalName
    }
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
ConnectToAzureAD

$global:ExportedUser = 0
$global:ReportTime = ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"

FindUseCase

if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") {     
    #Open file after code execution finishes
    Write-Host "The output file available in $global:ExportCSVFileName" -ForegroundColor Green 
    write-host "Exported $global:ExportedUser records to CSV." 
    $prompt = New-Object -ComObject wscript.shell    
    $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
    If ($userInput -eq 6) {    
        Invoke-Item "$global:ExportCSVFileName"
    }  
} else {
    #Notification when usecase doesn't have the data in the tenant
    Write-Host "No data found with the specified criteria"
}
Write-Host `nFor more Microsoft 365 reports"," please check o365reports.com -ForegroundColor Cyan
Disconnect-AzureAD
Write-host "`nDisconnected AzureAD Session Successfully"
