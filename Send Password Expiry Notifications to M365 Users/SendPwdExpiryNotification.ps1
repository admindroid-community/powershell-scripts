<#
=============================================================================================
Name:           Send Password Expiry Notifications to users in Microsoft 365
Version:        1.0
Website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Sends password expiry notifications to users about upcoming password expiry.
2. Filters results to display licensed users alone.
3. Export a list of users who fall under the given criteria
4. Automatically install the Microsoft Graph PowerShell module (if not installed already) upon your confirmation.
5. The script can be executed with MFA and non-MFA accounts.
6. It can be executed with certificate-based authentication (CBA) too.
7. The script is schedular-friendly – automatically schedule the script in the task scheduler to automate sending password expiry notifications.


For detailed Script execution:  https://o365reports.com/2025/02/11/send-password-expiry-notification-in-microsoft-365/
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $True)]
    [int]$DaysToExpiry,
    [Parameter(Mandatory = $false)]
    [switch]$LicensedUsersOnly,
    [switch]$Schedule,
    [string]$FromAddress,
    [string]$ClientId,
    [string]$TenantId,
    [string]$CertificateThumbprint,
    [Parameter(DontShow = $True)]
    [switch]$DoNotShowSummary
)

$Date = Get-Date
$CSVFilePath ="$(Get-Location)\PasswordExpiryNotificationSummary_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
[datetime]$StartTime = ($Date).AddDays(1).Date.AddHours(10)  #Update the value of AddHours to change the Schedule start time.
$ScriptPath = $MyInvocation.MyCommand.Path

# Function to connect to Microsoft Graph
function Connect_ToMgGraph {
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

    Write-Host "`nConnecting to Microsoft Graph..."
    
    if (($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne "")) {
        # Use certificate-based authentication if TenantId, ClientId, and CertificateThumbprint are provided
        Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    } else {
        # Use delegated permissions (Scopes) if credentials are not provided
        Connect-MgGraph -Scopes "User.Read.All", "Domain.Read.All", "Mail.Send.Shared" -NoWelcome 
    }

    # Verify connection
    if ((Get-MgContext) -ne $null) {
        if ((Get-MgContext).Account -ne $null) {
            Write-Host "Connected to Microsoft Graph PowerShell using account: $((Get-MgContext).Account)"
        }
        else {
            Write-Host "Connected to Microsoft Graph PowerShell using certificate-based authentication."
        }
    } else {
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
}

Connect_ToMgGraph
if ((Get-MgContext).Account -ne $null){ 
    if ([string]::IsNullOrEmpty($FromAddress)) {
        $FromAddress = (Get-MgContext).Account
    }
} else {
    if ([string]::IsNullOrEmpty($FromAddress)) {
        Write-Host "`nError: FromAddress is required when using certificate-based authentication." -ForegroundColor Red
        Exit
    }
} 

# Schedule Task
if ($Schedule.IsPresent) {
    Write-Host "Configuring scheduled task to send password expiry notification..."

    # Validate mandatory parameters for scheduling
    if (-not $TenantId -or -not $ClientId -or -not $CertificateThumbprint) {
        Write-Host "`nError: TenantId, ClientId, and CertificateThumbprint are mandatory for scheduling the script." -ForegroundColor Red
        Exit
    }

    # Define the action and trigger for the schedule to execute the script
    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File `"$($ScriptPath)`" -DoNotShowSummary -TenantId `"$($TenantId)`" -ClientId `"$($ClientId)`" -CertificateThumbprint `"$($CertificateThumbprint)`" -DaysToExpiry $($DaysToExpiry) -FromAddress `"$($FromAddress)`" -WindowStyle Hidden"
    $Trigger = New-ScheduledTaskTrigger -Daily -At $StartTime
    $Principal = New-ScheduledTaskPrincipal -UserId $env:UserName -LogonType Interactive -RunLevel Highest

    # Register the task
    $TaskName = "Password Expiry Notification"
    
    try {
        Register-ScheduledTask -Action $Action -Trigger $Trigger -Principal $Principal -TaskName $TaskName -Description "Runs the Password Expiry Notification script" -ErrorAction Stop | Out-Null
        Write-Host "`nScheduled task '$TaskName' created successfully and it will run daily at $($StartTime.ToString('hh:mm tt')) from $($StartTime.ToString('yyyy-MM-dd'))." -ForegroundColor Cyan
    }
    catch {
        Write-Host "`nError: Failed to register the scheduled task." -ForegroundColor Red
        Write-Host "Details: $_" -ForegroundColor Red
        Exit
    }
}

$Counter = 0
$PwdExpirigUsersCount = 0
$Domains = @{}

Get-MgDomain | Select-Object Id, AuthenticationType, PasswordValidityPeriodInDays | ForEach-Object {
    if ($_.AuthenticationType -eq "Federated") { 
        $Domains[$_.Id] = 0 
    } 
    else { 
        if ($_.PasswordValidityPeriodInDays -ne $null) { 
            $Domains[$_.Id] = $_.PasswordValidityPeriodInDays 
        } else { 
            $Domains[$_.Id] = 90
        }
    }
}
Get-MgUser -Filter "accountEnabled eq true" -All -Property AssignedLicenses, PasswordPolicies, DisplayName, UserPrincipalName, LastPasswordChangeDateTime | Where-Object {$_.PasswordPolicies -notcontains "DisablePasswordExpiration"} | ForEach-Object {
    $Name = $_.DisplayName
    $EmailAddress = $_.UserPrincipalName
    $LicenseStatus = $_.AssignedLicenses
    $PwdLastChangeDate = $_.LastPasswordChangeDateTime
    $UserDomain = $EmailAddress.Split('@')[1]
    $MaxPwdAge = $Domains[$UserDomain]
    
    $Counter++
    Write-Progress -Activity "Processed Users: $($Counter)" -Status "Processing $($_.DisplayName)"

    if ($MaxPwdAge -ne 2147483647) { 
        $ExpiryDate = $PwdLastChangeDate.AddDays($MaxPwdAge)
        $DaysToExpire = ($ExpiryDate.Date - $Date.Date).Days
    
        if ($LicenseStatus -ne $null) { $LicenseStatus = "Licensed" } else { $LicenseStatus = "Unlicensed" }
    
        if($DaysToExpire -eq 0){
            $Msg = "Today"
        }
        elseif($DaysToExpire -eq 1){
            $Msg = "Tomorrow"
        }
        else{
            $Msg = "in " + "$DaysToExpire" + " days"
        }
    
        $params = @{
	        message = @{
		        subject = "Your Password is About to Expire – Update Required"
		        body = @{
			        contentType = "HTML"
			        content = "Hello <b>$($Name)</b>,
                               <p>Your Microsoft 365 account password will expire <b><i>$($Msg)</i></b>.</p>
                               <p>To avoid disruptions in accessing Microsoft 365 apps and services, please update your password promptly by visiting the <a href=`"https://mysignins.microsoft.com/security-info/password/change`" target=`"_blank`">secure Microsoft Office portal</a>.</p>
                               <p>If you have any questions or encounter issues, feel free to contact help desk for assistance.</p>
                               <p>Thank you for your attention to this matter.</p>
                               Best regards,<br>  
                               IT Admin Team."
		        }
		        toRecipients = @(
			        @{
				        emailAddress = @{
					        address = $EmailAddress
				        }
			        }
		        )
	        }
        }

        [switch] $ProcessUser = $True

        if (($LicensedUsersOnly.IsPresent) -and ($LicenseStatus -ne "Licensed")) { $ProcessUser = $false }

        if ($ProcessUser.IsPresent) {
            if (($DaysToExpire -ge "0") -and ($DaysToExpire -le $DaysToExpiry)){
                $PwdExpirigUsersCount++
                Send-MgUserMail -UserId $FromAddress -BodyParameter $params 
                $ExportResult = @{'Name' = $Name; 'Email Address' = $EmailAddress; 'Days to Expire' = $DaysToExpire; 'Password Expiry Date' = $ExpiryDate; 'License Status' = $LicenseStatus}
                $ExportResults = New-Object PSObject -Property $ExportResult
                $ExportResults | Select-object 'Name', 'Email Address', 'Days to Expire', 'Password Expiry Date', 'License Status' | Export-csv -path $CSVFilePath -NoType -Append -Force
            }
        }
    }
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1900+ Microsoft 365 reports. ~~" -ForegroundColor Green

if (!$DoNotShowSummary.IsPresent) {
    if ((Test-Path -Path $CSVFilePath) -eq "True") {     
        Write-Host `n"$PwdExpirigUsersCount users' password is about to expire in $DaysToExpiry days, view there details in: " -NoNewline -ForegroundColor Yellow
        Write-Host $CSVFilePath
        $prompt = New-Object -ComObject wscript.shell    
        $userInput = $prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)    
        if ($userInput -eq 6) {    
            Invoke-Item "$CSVFilePath"
        }  
    }
    else{
        Write-Host `n"No user(s) found with passwords expiring in the next $DaysToExpiry days."
    }
}