<#
=============================================================================================
Name:           Assign manager to Office 365 users based on the users' properties
Description:    This script assigns manager to Office 365 users based on the users' properties
Version:        1.0
Website:        m365scripts.com

Script Highlights:
1. The script uses MS Graph PowerShell and installs MS Graph PowerShell SDK (if not installed already) upon your confirmation. 
2. It can be executed with certificate-based authentication (CBA) too.
3. Assigns Manager in Office 365 by using more than 10+ user properties, such as filtering by department, job title, and city.
4. Furthermore, to assign a manager on a highly-filtered basis. You can use the following parameters.
     -ExistingManager – Overrides your existing manager.
     -ImportUsersFromCsvPath – Assign a manager to the bulk users through the CSV input file.
     -ProcessOnlyUnmanagedUsers – Assign a manager to unmanaged users in the specific user property.
     -GetAllUnmanagedUsers – Assign a manager to all unmanaged users.
5. Automatically, downloads a CSV file. The CSV file contains the usernames that match the given condition.
6. Credentials are passed as parameters, so worry not!
7. Generates a log file that contains the result status of your manager assignment.

For detailed script execution: https://m365scripts.com/microsoft365/set-up-manager-for-office-365-users-based-on-the-users-property
============================================================================================
#>
param (
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [string] $Properties =$null,
    [string] $ExistingManager=$null,
    [switch] $ProcessOnlyUnmanagedUsers,
    [string] $ImportUsersFromCsvPath=$null,
    [switch] $GetAllUnmanagedUsers,
    [string] $ManagerId=""
)
Function ConnectMgGraphModule
{
    $MsGraphBetaModule =  Get-Module Microsoft.Graph.Beta -ListAvailable
    if($MsGraphBetaModule -eq $null)
    { 
        Write-host "Important: Microsoft Graph Beta module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install Microsoft Graph Beta module? [Y] Yes [N] No  
        if($confirm -match "[yY]") 
        { 
            Write-host "Installing Microsoft Graph Beta module..."
            Install-Module Microsoft.Graph.Beta -Scope CurrentUser -AllowClobber
            Write-host "Microsoft Graph Beta module is installed in the machine successfully" -ForegroundColor Magenta 
        } 
        else
        { 
            Write-host "Exiting. `nNote: Microsoft Graph Beta module must be available in your system to run the script" -ForegroundColor Red
            Exit 
        } 
    }
    try{
        if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
        {  
            Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction SilentlyContinue -ErrorVariable ConnectionError|Out-Null
            if($ConnectionError -ne $null)
            {    
                Write-Host $ConnectionError -Foregroundcolor Red
                Exit
            }
        }
        else
        {
            Disconnect-MgGraph -ErrorAction SilentlyContinue| Out-Null
            Connect-MgGraph -Scopes "Directory.ReadWrite.All"  -ErrorAction SilentlyContinue -Errorvariable ConnectionError |Out-Null
            if($ConnectionError -ne $null)
            {
                Write-Host "$ConnectionError" -Foregroundcolor Red
                Exit
            }
        }
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Exit
    }
    Write-Host "Microsoft Graph Beta Powershell module is connected successfully`n" -ForegroundColor Green
}
Function RemoveManagedUsers {
    $Users1 = $Users
    $Global:Users = @()
    Foreach ($User in $Users1) {
        $CheckManager = $User.Manager.AdditionalProperties.displayName
        $Percent = $Count / $Users1.length * 100
        $Count++
        Write-Progress -Activity "Checking for user as they already having manager or not" -PercentComplete $Percent 
        if($CheckManager.length -eq 0){
             $Global:Users += $User
        }
    }
    Write-Progress -Activity "users" -Status "Ready" -Completed
}

Function AssignManager {
    if($ProcessOnlyUnmanagedUsers.IsPresent){
        RemoveManagedUsers
    }
    if(($global:Users).length -eq 0) {
        $log = "No User found for this Filter Criteria"
        Write-Warning($log)
        CloseConnection
    }
    if($global:AlreadyFromCSV -eq $false){
        ExportUsers
    }
    While ($true) {
        if($ManagerId -eq  ""){
            $ManagerId = Read-Host "Enter manager's UserPrincipalName or Objectid"
        }
        $Manager = $UsersList |Where-Object{$_.UserPrincipalName -eq $ManagerId -or $_.Id -eq $ManagerId}
        if($Manager.length -eq 0){
            Write-Warning "Enter the valid UserPrincipalName or object id"
            $ManagerId = ""
            continue
        }
        else {
            break
        }
    }
    $ErrorCount = 0
    Foreach ($User in $global:Users) {
        $log = "Adding $($Manager.DisplayName) to $($User.DisplayName)"
        $log>>$logfile
        $Percentage = $Count/$global:Users.length * 100
        Write-Progress "Assigning manager($($Manager.DisplayName)) to the user: $($User.UserPrincipalName) Processed Users Count: $Count" -PercentComplete $Percentage
        $Param = @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$($Manager.Id)"}
        Set-MgBetaUserManagerByRef -UserId $User.Id -BodyParameter $Param -ErrorAction SilentlyContinue -ErrorVariable Err 
        if($Err -ne $null)
        {
            $log = "Manager assignment failed"
            $log>>$logfile
            $ErrorCount++
            continue
        }
        $log = "Manager assigned successfully"
        $log>>$logfile
        $Count++
    }
    if($ErrorCount -ne $Users.Count)
    {
        Write-Host "The Manager($($Manager.DisplayName)) was assigned to your users Successfully"  -ForegroundColor Green
    }
    Write-Host "log file location $logfile"
    $prompt = New-Object -ComObject wscript.shell    
    $UserInput = $prompt.popup("Do you want to open Log file?", 0, "Open Output File", 4)    
    if ($UserInput -eq 6) {    
        Invoke-Item "$logfile"
    } 
    CloseConnection
}

Function ExportUsers {
        $Holders = @()
        $HeadName = 'UserName'
        Foreach($User in $global:Users){
            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name $HeadName -Value $User.UserPrincipalName
            $Holders += $Obj
        }
        $File = "ManagerAssignedUser"+$ReportTime+".csv"
        $Holders | Export-csv $File -NoTypeinformation
        Write-Host "Exported users are in the File Location-  $Path\$File "  -ForegroundColor Green
}
Function GetFilteredUsers{
    Foreach($Property in $FilteredProperties.Keys){
        $FilterProperty = $Property
        $FilterValue = "$($FilteredProperties[$Property])"
        $UsersList = $UsersList |?{$_.$FilterProperty -eq $FilterValue}
    }
    $global:Users = $UsersList
}

Function ExistingManager {
    $TargetManagerDetails = $UsersList |Where-Object{$_.UserPrincipalName -eq $ExistingManager -or $_.Id -eq $ExistingManager}
    $UsersList| Foreach {
        $Name = $_.Manager.AdditionalProperties.userPrincipalName
        if ($Name.length -ne 0) {
            Write-Progress -Activity "checking users having manager as $($TargetManagerDetails.DisplayName)" -Status "Processing : $Count - $($_.DisplayName)"   
            if(($Name).compareto($TargetManagerDetails.UserPrincipalName) -eq 0){
                $global:Users += $_
            }
            $Count++
        }
    }
}
Function ImportUsers {
    $UserNames = @()
    $global:AlreadyFromCSV = $true
    try
    {
        (Import-CSV -path $ImportUsersFromCsvPath) | #file must having header Username and their values as userprincipalname or objectid
        ForEach-Object {
            $UserNames += $_.Username
        }
    }
    catch
    {
        Write-Host $_.Exception.Message  -ForegroundColor Red
        CloseConnection
    }
    if($UserNames.length -eq 0) {
        Write-Warning "No usernames found at the csv file,located at $Path"
        CloseConnection
    }
    Foreach ($UserName in $UserNames) {
        $Global:Users += $UsersList |Where-Object{$_.UserPrincipalName -eq $UserName -or $_.Id -eq $UserName}
        Write-Progress "Retrieving user information from CSV file ,retrieved users count $Count" -Activity "users" -PercentComplete $Count
        $Count++
    }
    Write-Progress -Activity "users" -Status "Ready" -Completed
 }
Function AllUnManagedUsers {
    foreach($User in $UsersList){
        $GetManager = $User.Manager.AdditionalProperties
        Write-Progress "Retrieving user with no manager - users count $Count" -Activity "users" -PercentComplete $count
        $Count++
        if($getmanager.Count -eq 0)
        {
            $Global:Users += $User
        }
     }
     Write-Progress -Activity "users" -Status "Ready" -Completed
 }

Function GetFilterProperties{
    $FilteredProperties = @{}
    if($Properties -ne "")
    {
        $Properties= $Properties.Split(",")
        Foreach($Property in $Properties)
        {
            $PropertyExists = $UsersList | Get-Member|?{$_.Name -contains "$Property"}
            if($PropertyExists -eq $null)
            {
                Write-Host "$Property property is not available. Please provide valid property." -ForegroundColor Red
                CloseConnection
            }
            While($true)
            {
                $PropertyValue = Read-Host "Enter the $Property value"
                if($PropertyValue.Length -eq 0)
                {
                    Write-Host "Value couldn't be null. Please enter again." -ForegroundColor Red
                    Continue
                }
                break
            }
            $FilteredProperties.Add($Property,$PropertyValue)
        }
    }
    else
    {
        #if you want to add any property do it here at $userProperties (note: add only valid property with spellcheck)
        $UserProperties = @("","Department","JobTitle","CompanyName","City","Country","State","UsageLocation","UserPrincipalName","DisplayName","AgeGroup","UserType")
        if($Properties -eq ""){
            for ($index=1;$index -lt $UserProperties.length;$index++) {
                Write-Host("$index)$($UserProperties[$index])") -ForegroundColor Yellow
            }
            Write-Host "`nEnter your choice from 1 to $(($UserProperties.length)-1). If you want to filter users by multiple attributes, give them as comma separated value."
            Write-Host "(For example, if you want to filter users by Department and State, enter your choice as 5,8)" -ForegroundColor Yellow
            [string]$Properties = Read-Host("Enter your choice")
        }
        while($Properties -eq "")
        {
            Write-Host "Choice couldn't be null. Please enter again." -ForegroundColor Red
            [string]$Properties = Read-Host "`nEnter your choice"
        }
        try{
            [int []]$choice = $Properties.split(',')
            for($i=0;$i -lt $choice.Length;$i++){
                [int]$index = $choice[$i]
                $propertyValue = Read-Host "Enter $($UserProperties[$index]) value"
                if(($propertyValue.length -eq 0)){
                    Write-Host "Value couldn't be null. Please enter again." -ForegroundColor Red
                    $i--;
                    continue
                }
                $FilteredProperties.Add($UserProperties[$index],$propertyValue)
            }
        }
        catch{
            Write-Host $_.Exception.Message -ForegroundColor Red
            CloseConnection
        }
    }
    GetFilteredUsers
}
Function CloseConnection
{
    Disconnect-MgGraph | Out-Null
    Write-Host "Session disconnected successfully" 
    Exit
}
ConnectMgGraphModule
Write-Host "`nNote: If you encounter module related conflicts, run the script in a fresh Powershell window." -ForegroundColor Yellow
$UsersList = Get-MgBetaUser -All -ExpandProperty Manager
$Global:Users = @()
$Count = 1
$ReportTime = ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString())
$LogFileName = "LOGfileForManagerAssignedUser"+$ReportTime+".txt"
$path = (Get-Location).path
$logfile = "$path\$LogFileName"
$global:AlreadyFromCSV = $false
#.................only param properties...................
if($ExistingManager.Length -ne 0){
    ExistingManager
    AssignManager
}
if($ImportUsersFromCsvPath.Length -ne 0){
    ImportUsers
    AssignManager
}
if($GetAllUnmanagedUsers.IsPresent){
    AllUnManagedUsers
    AssignManager
}
#.........................................................
GetFilterProperties
AssignManager

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n