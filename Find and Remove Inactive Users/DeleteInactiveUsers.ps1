<#
======================================================================================================
Name: Identify and Remove Inactive Users in Microsoft 365
Version: 1.0
Website: admindroid.com

Script Highlights:
1. The script automatically verifies and installs the Microsoft Graph PowerShell SDK module (if not installed already) upon your confirmation.
2. Generates and exports all the inactive Microsoft 365 users into a CSV file.
3. Identifies sign-in enabled inactive users and disable their account.
4. Retrieves all the licensed inactive users and deletes them to reuse licenses.
5. Finds all the sign-in disabled users and deletes their accounts.
6. Identifies the external inactive users and removes them from the organization.
7. Allows to use the previously generated inactive users report and take actions later (i.e disable or delete).
8. The script is scheduler-friendly.
9. Supports certificate-based authentication (CBA) too.

For detailed Script execution: https://blog.admindroid.com/identify-and-remove-inactive-users-in-microsoft-365/

============================================================================================================
#>






Param
(
    [int]$InactiveDays,
    [int]$InactiveDays_NonInteractive,
    [ValidateSet('Delete','Disable')]
    [string]$Action,
    [switch]$GenerateReportOnly,
    [switch]$ExcludeNeverLoggedInUsers,
    [switch]$EnabledUsersOnly,
    [switch]$DisabledUsersOnly,
    [switch]$LicensedUsersOnly,
    [switch]$ExternalUsersOnly,
    [switch]$Force,
    [switch]$CreateSession,
    [string]$ImportCsv,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

# Function to connect to Microsoft Graph
Function Connect_MgGraph {
    # Check if Microsoft Graph module is installed
    $MsGraphModule = Get-Module Microsoft.Graph -ListAvailable
    if ($MsGraphModule -eq $null) {
        Write-Host "`nImportant: Microsoft Graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host "Are you sure you want to install Microsoft Graph module? [Y] Yes [N] No"
        if ($confirm -match "[yY]") {
            Write-Host "Installing Microsoft Graph module..."
            Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber
            Write-Host "Microsoft Graph module is installed in the machine successfully" -ForegroundColor Magenta 
        } else {
            Write-Host "Exiting. `nNote: Microsoft Graph module must be available in your system to run the script" -ForegroundColor Red
            Exit
        }
    }
    
    # Disconnects existing connection
    if($CreateSession.IsPresent)
    {
        Disconnect-MgGraph -WarningAction SilentlyContinue | Out-Null
    }

    Write-Host "`nConnecting to Microsoft Graph..."
    if (($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne "")) {
        # Use certificate-based authentication if TenantId, ClientId, and CertificateThumbprint are provided
        Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    } else {
        # Use delegated permissions (Scopes) if credentials are not provided
        Connect-MgGraph -Scopes "User.EnableDisableAccount.All","User.DeleteRestore.All" -NoWelcome 
    }
 
    # Verify connection
    if ((Get-MgContext) -ne $null) {
        Write-Host "Connected to Microsoft Graph PowerShell using account: $((Get-MgContext).Account)`n"
    } else {
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
}

Connect_MgGraph

# Function to validate mandatory parameters (inactive days & action)
Function mandatory-parameter-validation{
    # Inline validation for InactiveDays
    if ((!$InactiveDays -and !$InactiveDays_NonInteractive) -and ($ImportCsv -eq "")) {
        [int]$Script:InactiveDays = Read-Host "`n Enter InactiveDays"
        if ([string]::IsNullOrWhiteSpace($InactiveDays) -or $InactiveDays -le 0) {
            Write-Host "`n Invalid InactiveDays provided. Exiting script." -ForegroundColor Red
            Exit
        }
    }

    # Inline validation for Action
    if (!$GenerateReportOnly -and ($Action -eq "" )) {
        $Script:Action = Read-Host "`n Enter the Action (Delete or Disable)"
        if ($Action -ne 'Delete' -and $Action -ne 'Disable') {
            Write-Host "`n Invalid Action provided. Exiting script." -ForegroundColor Red
            Exit
        }
    }
}

# Function to handle deletion/disabling
Function Delete_Inactive_Users {
    param($UserRecord)

    $LogEntry = @()
    $User_Id = $UserRecord.UPN
    $Acc_Status = $UserRecord.'Account Status'
    Try{
        switch ($Action) {
            'Delete' {
                Remove-MgUser -UserId $User_Id -ErrorAction Stop
                $LogEntry += "[INFO] - '$User_Id' is deleted."
                break
             }
            'Disable' {
                if($Acc_Status -eq 'Disabled'){
                    $LogEntry += "[INFO] - '$User_Id' is disabled already."
                }else{
                    Update-MgUser -UserId $User_Id -AccountEnabled:$false -ErrorAction Stop
                    $LogEntry += "[INFO] - '$User_Id' is disabled."
                }
            break
            }
        }
    }catch{
        $LogEntry += "[ERROR] - Failed to $Action the user '$User_Id': $($_.Exception.Message)`n"
    }

   
        $LogEntry | Out-File -FilePath $LogFilePath -Append
  
}

# Fuction to handle user confirmation to perform deletion operation
Function Confirm_User_Deletion_Action{
    param($CsvFile)

    # Automatically performs action without confirmation
    if ($Force) {
        Import-Csv -Path $CsvFile | ForEach-Object{ Delete_Inactive_Users -UserRecord $_ }
        Write-Host "`n The log file is available in: " -NoNewline -ForegroundColor Yellow
        Write-Host " $LogFilePath"
    } 
    # Asks confirmation to perform operation
    else {
        $Confirm = $(Write-Host "`n Are you sure you want to $Action $PrintedUser inactive users in the CSV file? [Y/N]: " -ForegroundColor Yellow -NoNewline; Read-Host)
        if ($Confirm -match "[yY]") {
            Import-Csv -Path $CsvFile | ForEach-Object { Delete_Inactive_Users -UserRecord $_ }
            Write-Host "`n The log file is available in: " -NoNewline -ForegroundColor Yellow
            Write-Host " $LogFilePath"
            $Prompt = New-Object -ComObject wscript.shell
            $UserInput = $Prompt.popup("Do you want to open the log file?", 0, "Open Output File", 4)
            if ($UserInput -eq 6) {
                Invoke-Item "$LogFilePath"
            }
        }else{
            Write-Host "`n No action performed."
        }
    }
}

# Function to filter inactive users from all the available users
Function Filter_Inactive_Users{

    # Process each user
    $RequiredProperties=@('UserPrincipalName','EmployeeId','CreatedDateTime','AccountEnabled','Department','JobTitle','RefreshTokensValidFromDateTime','SigninActivity') 
    Get-MgUser -All -Property $RequiredProperties | Select $RequiredProperties | ForEach-Object {
    $Count++
    $UPN = $_.UserPrincipalName
    Write-Progress -Activity "`nProcessing user: $Count - $UPN"
    $EmployeeId = $_.EmployeeId
    $LastInteractiveSignIn = $_.SignInActivity.LastSignInDateTime
    $LastNon_InteractiveSignIn = $_.SignInActivity.LastNonInteractiveSignInDateTime
    $CreatedDate = $_.CreatedDateTime
    $AccountEnabled = $_.AccountEnabled
    $Department = $_.Department
    $JobTitle = $_.JobTitle

    # Calculate inactive days for interactive sign-ins
    if($LastInteractiveSignIn -eq $null)
     {
      $LastInteractiveSignIn = "Never Logged In"
      $InactiveDays_InteractiveSignIn = "-"
     }
     else
     {
      $InactiveDays_InteractiveSignIn = (New-TimeSpan -Start $LastInteractiveSignIn).Days
     }
     if($LastNon_InteractiveSignIn -eq $null)
     {
      $LastNon_InteractiveSignIn = "Never Logged In"
      $InactiveDays_NonInteractiveSignIn = "-"
     }
     else
     {
      $InactiveDays_NonInteractiveSignIn = (New-TimeSpan -Start $LastNon_InteractiveSignIn).Days
     }

     # Get user account status
     if($AccountEnabled -eq $true)
     {
      $AccountStatus='Enabled'
     }
     else
     {
      $AccountStatus='Disabled'
     }

     #Get licenses assigned to mailboxes
     $Subscriptions = Get-MgUserLicenseDetail -UserId $UPN | Select SkuId, SkuPartNumber
     $Licenses = $Subscriptions.SkuPartNumber
     $AssignedLicense = @()

     #Convert license plan to friendly name
     if($Licenses.count -eq 0)
     {
      $LicenseDetails = "No License Assigned"
     }
     else
     {
      foreach($License in $Licenses)
      {
       $EasyName = $FriendlyNameHash[$License]
       if(!($EasyName))
       {$NamePrint = $License}
       else
       {$NamePrint = $EasyName}
       $AssignedLicense += $NamePrint
      }
      $LicenseDetails = $AssignedLicense -join ", "
     }
     $Print = 1

     #Inactive days based on interactive signins filter
     if($InactiveDays_InteractiveSignIn -ne "-"){
      if(($InactiveDays -ne "") -and ($InactiveDays -gt $InactiveDays_InteractiveSignIn))
      {
       $Print=0
      }
     }
    
     #Inactive days based on non-interactive signins filter
     if($InactiveDays_NonInteractiveSignIn -ne "-"){
      if(($InactiveDays_NonInteractive -ne "") -and ($InactiveDays_NonInteractive -gt $InactiveDays_NonInteractiveSignIn))
      {
       $Print=0
      }
     }

    # Exclude never logged-in users
    if ($ExcludeNeverLoggedInUsers -and ($LastInteractiveSignIn -eq "Never Logged In")) {
        $Print = 0
    }

    # Filter for external users
    if ($ExternalUsersOnly -and $UPN -notmatch '#EXT#') {
        $Print = 0
    }

    #Signin Allowed Users
    if($EnabledUsersOnly.IsPresent -and $AccountStatus -eq 'Disabled'){      
        $Print=0
    }

    #Signin disabled users
    if($DisabledUsersOnly.IsPresent -and $AccountStatus -eq 'Enabled'){
        $Print=0
    }

    # Licensed users only filter
    if ($LicensedUsersOnly -and $Licenses.Count -eq 0){
        $Print = 0
    }

    # Generate report only
    if ($Print -eq 1) {
        $Script:PrintedUser++
        $ExportResult = [PSCustomObject]@{'UPN'=$UPN;'Last Interactive SignIn Date'=$LastInteractiveSignIn;'Last Non Interactive SignIn Date'=$LastNon_InteractiveSignIn;'Inactive Days(Interactive SignIn)'=$InactiveDays_InteractiveSignIn;'Inactive Days(Non-Interactive Signin)'=$InactiveDays_NonInteractiveSignIn;'License Details'=$LicenseDetails;'Account Status'=$AccountStatus;'Creation Date'=$CreatedDate;'Emp id'=$EmployeeId;'Department'=$Department;'Job Title'=$JobTitle}
        $ExportResult | Export-Csv -Path $ExportCSV -NoTypeInformation -Append
    } 
  }
}

$Location = Get-Location

$ExportCSV = "$Location\Inactive_M365_User_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$LogFilePath = "$Location\User_$($Action)_Log_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).txt"

$ExportResult = ""
$ExportResults = @()

$FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData

$Count=0
$PrintedUser=0

mandatory-parameter-validation

if ($ImportCsv -ne "") {
    $PrintedUser = Import-Csv -Path $ImportCsv | measure | Select -ExpandProperty Count
    if($PrintedUser -eq 0){
        Write-Host "`n No users found in $ImportCsv."
    }
    else{
        Confirm_User_Deletion_Action -CsvFile $ImportCsv
    }
}else {
    Filter_Inactive_Users
    if($PrintedUser -eq 0){
        Write-Host "`n Inactive users not found."
    }
    else{
        if(!$GenerateReportOnly){
            Write-Host "`n Detailed inactive users report available in: " -ForegroundColor Yellow
            Write-Host "`n $ExportCSV"
            Confirm_User_Deletion_Action -CsvFile $ExportCSV
        }
    }
}

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n `n

if($GenerateReportOnly.IsPresent){
    if((Test-Path -Path $ExportCSV) -eq "True")
    {
        Write-Host "`n Detailed inactive users report available in: " -ForegroundColor Yellow
        Write-Host "`n $ExportCSV"
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open the output file?", 0, "Open Output File", 4)
        if ($UserInput -eq 6) {
            Invoke-Item "$ExportCSV"
        }
    }
    else{
        Write-Host "`n No user found for the specific criteria" -ForegroundColor Red
    }
}