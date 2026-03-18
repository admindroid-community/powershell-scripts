<#
=============================================================================================
Name:           Manage User Deletions in Microsoft 365 Using PowerShell
Version:        1.0
website:        m365scripts.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~

1. Performs 8 user management actions and 1 reporting action related to user deletion. 
2. Automatically install the Microsoft Graph PowerShell module (if not already installed) after your confirmation. 
3. Supports execution with MFA-enabled accounts. 
4. Exports deleted users report results to a CSV file. 
5. Tracks the execution status of all deletion actions and exports them to a CSV file. 
6. Supports certificate-based authentication. 
7. The script is scheduler friendly. 

For detailed script execution: https://m365scripts.com/microsoft365/manage-user-deletions-in-microsoft-365-using-powershell/
============================================================================================
#>

Param (
    [string]$Action = "0",
    [switch]$MultiExecutionMode,
    [string]$InputCsvFilePath,
    [array]$UPN,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Function Connect_MgGraph {
    #Check for module installation
    $Module = Get-Module -Name microsoft.graph -ListAvailable
    if($Module.Count -eq 0){
        Write-Host "Microsoft Graph PowerShell SDK is not available"  -ForegroundColor yellow 
        $Confirm = Read-Host Are you sure want to install the module? [Y]Yes [N]No
        if($Confirm -match "[Yy]"){
            Write-Host "Installing Microsoft Graph PowerShell module..."
            Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
        }
        else{
            Write-Host "Microsoft Graph PowerShell module is required to run this script. Please install module using 'Install-Module Microsoft.Graph' cmdlet." 
            Exit
        }
    }
    #Disconnect Existing MgGraph session
    if ($CreateSession.IsPresent) {
    	Disconnect-MgGraph | Out-Null
    }
     Write-Host "`nConnecting to Microsoft Graph..." 
    if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
    {  
        Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    }
    else{
        Connect-MgGraph -Scopes "User.DeleteRestore.All", "AuditLog.Read.All", "User.Read.All" -NoWelcome
    }
}

function Log-UserDeletion {
    param (
        [string]$UserId,
        [string]$Operation,
        [boolean]$Status,
        [string]$ErrorMessage
    )

    $script:LatestFile = $LogFilePath
    $script:LastOutputType = "Log"
    $Timestamp = (Get-Date).ToString('yyyy-MM-dd_HH-mm-ss')
    $CSVEntry = [PSCustomObject]@{
        "Event Time"        = $Timestamp
        "UserPrincipalName" = $UserId
        "Operation"         = $Operation
        "Status"            = if ($Status) { "Success" } else { "Failed" }
        "Error Message"     = if ([string]::IsNullOrEmpty($ErrorMessage)) { "-" } else { $ErrorMessage }
    }

    $CSVEntry | Export-Csv -Path $LogFilePath -NoTypeInformation -Append -Force

    if ($script:SingleOperation) {
        $script:OperationStatus = $Status
        if ($Status) {
            $script:OperationMessages += "$Operation $UserId succeeded."
        }
        else {
            $script:OperationMessages += "$Operation $UserId failed. Error: $ErrorMessage"
        }
    }
}

function Get-UPNsFromCsv {
    param([string]$CsvPath)

    if ([string]::IsNullOrWhiteSpace($CsvPath)) {
        $CsvPath = Read-Host "Enter CSV file path"
    }
    if (!(Test-Path $CsvPath)) {
        Write-Host "Invalid CSV file path." -ForegroundColor Red
        return $null
    }
    $csv = Import-Csv $CsvPath
    if ($csv.Count -eq 0) {
        Write-Host "CSV file is empty." -ForegroundColor Red
        return $null
    }
    if (-not $csv[0].PSObject.Properties.Name.Contains("UserPrincipalName")) {
        Write-Host "CSV file is missing 'UserPrincipalName' column." -ForegroundColor Red
        return $null
    }
    return $csv.UserPrincipalName
}

function Exit-Script {
    param([switch]$Exit)

    if($script:LatestFile -and (Test-Path $script:LatestFile)){
        if ($script:LastOutputType -eq "Report") {
            Write-Host "`nThe report file is available in: " -NoNewline -ForegroundColor Yellow; Write-Host $script:LatestFile;
        }
        elseif ($script:LastOutputType -eq "Log") {
            Write-Host "`nThe log file is available in: " -NoNewline -ForegroundColor Yellow; Write-Host $script:LatestFile;
        }
    }
    $script:ExecutionCount = 0
    Disconnect-MgGraph | Out-Null
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green 
    if ($script:LatestFile){
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open the file?",0,"Open File",4)
        if ($UserInput -eq 6) {
            Invoke-Item $script:LatestFile
        }
    }
    if ($Exit.IsPresent) { Exit 0 }
}

function Delete-Users {
    param (
        [Parameter(Mandatory=$true)]
        [array]$UserPrincipalNames,
        [switch]$HardDelete
    )

    foreach ($user in $UserPrincipalNames) {
        $targetId = $null
        try {
            $activeUser = Get-MgUser -Filter "UserPrincipalName eq '$user'" -Property "Id" -ErrorAction Stop
            
            if ($activeUser) {
                $targetId = $activeUser.Id
                Remove-MgUser -UserId $targetId -Confirm:$false
            } 
            else {
                $deletedUser = Get-MgDirectoryDeletedItemAsUser -Filter "UserPrincipalName eq '$user'" -Property "Id" -ErrorAction Stop
                
                if ($deletedUser) {
                    $targetId = $deletedUser.Id
                }
            }

            if ($HardDelete.IsPresent -and $targetId) {
                Start-Sleep -Seconds 1
                Remove-MgDirectoryDeletedItem -DirectoryObjectId $targetId -Confirm:$false
                
                Log-UserDeletion -UserId $user -Operation "Permanently delete user" -Status $true -ErrorMessage "-"
            }
            elseif (!$targetId) {
                Log-UserDeletion -UserId $user -Operation "Delete user" -Status $false -ErrorMessage "User not found"
            }
            elseif (!$HardDelete.IsPresent) {
                Log-UserDeletion -UserId $user -Operation "Delete user" -Status $true -ErrorMessage "-"
            }

        } 
        catch {
            Log-UserDeletion -UserId $user -Operation "Delete user" -Status $false -ErrorMessage $_.Exception.Message
        }
    }
}

function Restore-DeletedUsers {
    param (
        [Parameter(Mandatory=$true)]
        [array]$UPNs
    )

    foreach ($UPN in $UPNs) {
        try {
            $match = Get-MgDirectoryDeletedItemAsUser -Filter "userPrincipalName eq '$UPN'" -Property "Id" -ErrorAction Stop

            if ($match) {
                Restore-MgDirectoryDeletedItem -DirectoryObjectId $match.Id -ErrorAction Stop | Out-Null
                Log-UserDeletion -UserId $UPN -Operation "Restore user" -Status $true
            } 
            else {
                Log-UserDeletion -UserId $UPN -Operation "Restore user" -Status $false -ErrorMessage "User not found in deleted items."
            }
        }
        catch {
            Log-UserDeletion -UserId $UPN -Operation "Restore user" -Status $false -ErrorMessage $_.Exception.Message
        }
    }
}

function Restore-AllDeletedUsers {

    try {
        $count = 0
        Get-MgDirectoryDeletedItemAsUser -All -Property "Id,UserPrincipalName" -ErrorAction Stop | ForEach-Object {
            $count++
            $UPN = $_.UserPrincipalName
            Write-Progress -Activity "Restoring deleted user "-Status "$($UPN)"

            try {
                Restore-MgDirectoryDeletedItem -DirectoryObjectId $_.Id -ErrorAction Stop | Out-Null
                Log-UserDeletion -UserId $UPN -Operation "Restore user" -Status $true
            }
            catch {
                Log-UserDeletion -UserId $UPN -Operation "Restore user" -Status $false -ErrorMessage $_.Exception.Message
            }
        }
        Write-Progress -Activity "Restoring deleted users..." -Completed
        if ($count -eq 0) {
            Write-Host "No deleted users found." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Permanently-Delete-DeletedUsers {
    try {
        $count = 0
        Get-MgDirectoryDeletedItemAsUser -All -Property "Id,UserPrincipalName" | ForEach-Object {
            $count++
            $UPN = $_.UserPrincipalName
            Write-Progress -Activity "Permanently deleting user"-Status "$($UPN)" -Completed

            try {
                Remove-MgDirectoryDeletedItem -DirectoryObjectId $_.Id -Confirm:$false -ErrorAction Stop
                Log-UserDeletion -UserId $UPN -Operation "Permanently delete user" -Status $true -ErrorMessage ""
            }
            catch {
                Log-UserDeletion -UserId $UPN -Operation "Permanently delete user" -Status $false -ErrorMessage $_.Exception.Message
            }
        }
        Write-Progress -Activity "Permanently deleting users..." -Completed
        if ($count -eq 0) {
            Write-Host "No deleted users found." -ForegroundColor Yellow
        }
    }
    catch {
        Log-UserDeletion -Operation "Permanently delete all deleted users" -Status $false -ErrorMessage $_.Exception.Message
    }
}

function List-DeletedUsers {
    try {
        $timestamp = $((Get-Date).ToString('yyyy-MM-dd-ddd hh-mm-ss tt'))
        $script:ExportPath = "$Location\M365_Deleted_Users_Report_$timestamp.csv"
        $script:LatestFile = $script:ExportPath; $script:LastOutputType = "Report"

        $properties = "id,displayName,userPrincipalName,lastPasswordChangeDateTime,accountEnabled,proxyAddresses,mail,userType,jobTitle,department,deletedDateTime,createdDateTime,signInActivity,creationType,companyName,mobilePhone"
        
        $count = 0

        Get-MgDirectoryDeletedItemAsUser -All -Property $properties | ForEach-Object {
            $count++
            $userData = [PSCustomObject]@{
                "Display Name"                      = $_.displayName
                "UserPrincipalName"               = $_.userPrincipalName
                "User Type"                         = if ($_.userType) { $_.userType } else { "-" }
                "Email Address"                     = if ($_.mail) { $_.mail } else { "-" }
                "Created Date Time"                 = $_.createdDateTime
                "Deleted Date Time"                 = $_.deletedDateTime
                "Last Successful Sign-In Date Time" = if ($_.signInActivity.lastSignInDateTime) { $_.signInActivity.lastSignInDateTime } else { "-" }
                "Last Password Change Date Time"    = $_.lastPasswordChangeDateTime
                "Proxy Addresses"                   = if ($_.proxyAddresses) { $_.proxyAddresses -join ", " } else { "-" }
                "User Id"                           = $_.id
                "Job Title"                         = if ($_.jobTitle) { $_.jobTitle } else { "-" }
                "Department"                        = if ($_.department) { $_.department } else { "-" }
                "Mobile Phone"                      = if ($_.mobilePhone) { $_.mobilePhone } else { "-" }
            }

            $userData | Export-Csv -Path $script:ExportPath -NoTypeInformation -Encoding UTF8 -Append
        }
        if ($count -eq 0) {
            Write-Host "`nNo deleted users found." -ForegroundColor Yellow
        }
        
        if ($MultiExecutionMode.IsPresent -and (Test-Path -Path $script:ExportPath)) {
            Write-Host "`nThe report file is available in: " -NoNewline -ForegroundColor Yellow; Write-Host "$script:ExportPath"
        }

    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

$Location = Get-Location
$script:Message = ""
$script:OperationMessages = @()
$script:OperationStatus = $null;
$script:LatestFile = $null; $script:LastOutputType = $null
$LogFilePath = "$Location\M365_User_Deletion_Management_Log_$((Get-Date).ToString('yyyy-MM-dd-ddd hh-mm-ss tt')).csv"

Connect_MgGraph
do {

    if ($Action -eq 0) {
        Write-Host "`n===================================================" -ForegroundColor Cyan
        Write-Host "     Microsoft 365 User Deletion Management" -ForegroundColor Green
        Write-Host "===================================================" -ForegroundColor Cyan
        Write-Host @"
        1. Delete user
        2. View all deleted users
        3. Permanently delete user
        4. Restore deleted user
        5. Permanently delete all soft deleted users
        6. Delete bulk users (CSV input)
        7. Permanently delete bulk users (CSV input)
        8. Bulk restore deleted users (CSV input)
        9. Restore all deleted users
       10. Exit
"@ -ForegroundColor Yellow
        Write-Host "`n===================================================" -ForegroundColor Cyan

        $Action = (Read-Host "`nPlease choose the action to continue").Trim()

    }
    if ($Action -in @(1,3,4)) {
        if ($UPN) {
            $UPNs = $UPN
        }
        else {
            $inputUPN = Read-Host "`nEnter the UserPrincipalName of the user to perform the action"
            $UPNs = @($inputUPN.Trim())
        }
        $script:SingleOperation = $true
    }

    switch ($Action) {
        "1"  { Delete-Users -UserPrincipalNames $UPNs }
        "2"  { List-DeletedUsers }
        "3"  { Delete-Users -UserPrincipalNames $UPNs -HardDelete }
        "4"  { Restore-DeletedUsers -UPNs $UPNs }
        "5"  { Permanently-Delete-DeletedUsers }
        "6"  { 
                $UPNs = Get-UPNsFromCsv $InputCsvFilePath
                if ($UPNs) {
                    Delete-Users -UserPrincipalNames $UPNs
                }
             }
        "7"  {
                $UPNs = Get-UPNsFromCsv $InputCsvFilePath
                if ($UPNs) {
                    Delete-Users -UserPrincipalNames $UPNs -HardDelete
                }
             }
        "8"  {
                $UPNs = Get-UPNsFromCsv $InputCsvFilePath
                if ($UPNs) {
                    Restore-DeletedUsers -UPNs $UPNs
                }
             }
        "9"  { Restore-AllDeletedUsers }
        "10" { Exit-Script -Exit }
        default {
            Write-Host "`nInvalid choice. Please select a valid action." -ForegroundColor Red
            if(!($MultiExecutionMode.IsPresent)) {Exit-Script -Exit}
        }
    }

    if ($MultiExecutionMode.IsPresent -and ($script:LastOutputType -eq "Log") -and ($Action -notin @(1,2,3,4)) -and (Test-Path $LogFilePath)) {
    Write-Host "`nThe log file is available in: " -NoNewline -ForegroundColor Yellow
    Write-Host $LogFilePath
    }

    if ($script:SingleOperation -and ($script:OperationMessages.Count -gt 0) -and ($Action -in @(1,3,4))) {
        Write-Host ""
        foreach ($msg in $script:OperationMessages) {
            if ($msg -match "succeeded") {
                Write-Host $msg -ForegroundColor Green
            }
            else {
                Write-Host $msg -ForegroundColor Red
            }
        }
        $script:OperationMessages = @()
    }

    $script:OperationStatus = $null; $script:Message = ""

    if ($MultiExecutionMode.IsPresent) { 
        if ($UPN) { $UPN.Clear() }
        $InputCsvFilePath = $null
        $UPNs = $null
        $Action = 0 
    }
} while ($MultiExecutionMode.IsPresent)

Exit-Script -Exit