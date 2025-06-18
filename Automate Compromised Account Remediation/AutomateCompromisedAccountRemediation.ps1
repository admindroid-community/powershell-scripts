<#
========================================================================================
Name:                    Automate Compromised Microsoft Account Remediation
Description:             This script automates compromised account remediation with 8 best practices
Version:                 1.0
Website:                 o365reports.com

Script Highlights:    
~~~~~~~~~~~~~~~~~~  
1.Automates 8 best practices of compromised user remediation. 
2.Automatically installs required PowerShell modules MS Graph and Exchange Online PowerShell (if not installed already) upon your confirmation.  
3.Export the detailed log file on actions performed and their status (success or failure). 
4.Supports certificate-based authentication (CBA) too.  

For detailed script execution: https://o365reports.com/2025/06/17/automate-compromised-account-remediation-microsoft-365/ 
=========================================================================================
#>
                  


param(
[string]$TenantId,
[string]$ClientId,
[string]$CertificateThumbprint,
[string]$CSVFilePath,
[String]$UPNs
)

Function ConnectModules 
{
    $MsGraphModule =  Get-Module Microsoft.Graph -ListAvailable
    if($MsGraphModule -eq $null)
    { 
        Write-host "Important: Microsoft Graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install Microsoft Graph module? [Y] Yes [N] No  
        if($confirm -match "[yY]") 
        { 
            Write-host "Installing Microsoft Graph module..."
            Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber
            Write-host "Microsoft Graph module is installed in the machine successfully" -ForegroundColor Magenta 
        } 
        else
        { 
            Write-host "Exiting. `nNote: Microsoft Graph module must be available in your system to run the script" -ForegroundColor Red
            Exit 
        } 
    }
    $ExchangeOnlineModule =  Get-Module ExchangeOnlineManagement -ListAvailable
    if($ExchangeOnlineModule -eq $null)
    { 
        Write-host "Important: Exchange Online module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install Exchange Online module? [Y] Yes [N] No  
        if($confirm -match "[yY]") 
        { 
            Write-host "Installing Exchange Online module..."
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
            Write-host "Exchange Online Module is installed in the machine successfully" -ForegroundColor Magenta 
        } 
        else
        { 
            Write-host "Exiting. `nNote: Exchange Online module must be available in your system to run the script" 
            Exit 
        } 
    }

    Write-Host "Connecting modules(Microsoft Graph and Exchange Online module)...`n"
    try{
        if($TenantId -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
        {
            Connect-MgGraph  -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction SilentlyContinue -ErrorVariable ConnectionError|Out-Null
            if($ConnectionError -ne $null)
            {    
                Write-Host $ConnectionError -Foregroundcolor Red
                Exit
            }
            Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization (Get-MgDomain | Where-Object {$_.isInitial}).Id -ShowBanner:$false
        }
        else
        {
            Connect-MgGraph -Scopes "Directory.AccessAsUser.All","user.ReadWrite.All","UserAuthenticationMethod.Read.All"  -ErrorAction SilentlyContinue -Errorvariable ConnectionError |Out-Null #Directory.ReadWrite.All,AppRoleAssignment.ReadWrite.All,User.EnableDisableAccount.All,Directory.AccessAsUser.All,RoleManagement.ReadWrite.Directory
            if($ConnectionError -ne $null)
            {
                Write-Host $ConnectionError -Foregroundcolor Red
                Exit
            }
            Connect-ExchangeOnline -UserPrincipalName (Get-MgContext).Account -ShowBanner:$false
        }
    }
    catch
    {
        Write-Host $_.Exception.message -ForegroundColor Red
        Exit
    }
    Write-Host "Microsoft Graph PowerShell module is connected successfully" -ForegroundColor Cyan
    Write-Host "Exchange Online module is connected successfully" -ForegroundColor Cyan
}

Function DisableUser
{
    try{
        Update-MgUser -UserId $UPN -AccountEnabled:$false -ErrorAction Stop
        $Script:DisableUserAction = "Success"
    }
    catch
    {
        $Script:DisableUserAction = "Failed"
        $ErrorLog = "$($UPN) - Disable User Action - "+$Error[0].Exception.Message
        $ErrorLog>>$ErrorsLogFile
    }
}

Function SignOutFromAllSessions
{   
 try
 {
  Revoke-MgUserSignInSession -UserId $UPN -ErrorAction Stop | Out-Null
  $Script:SignOutFromAllSessionsAction = "Success"
 }
 Catch
 {
  $Script:SignOutFromAllSessionsAction = "Failed"
  $ErrorLog = "$($UPN) - Revoke user sign-in sessions - "+$_.exception.message
  $ErrorLog>>$ErrorsLogFile
 }
}

Function ResetPasswordToRandom
{
 $SpecialChars = [char[]]'!@#$%^&*()_+-=[]{}|;:,.<>?'
 $AllChars = ((48..57) + (65..90) + (97..122) | ForEach-Object { [char]$_ }) + $SpecialChars

 # Generate random password
 $Password = -join ($AllChars | Get-Random -Count 8)
 $log = "$UPN - $Password"
 $Pwd = ConvertTo-SecureString $Password -AsPlainText –Force
 try
 {
  $Passwordprofile = @{
   forceChangePasswordNextSignIn = $true
   password = $Password
  }
  Update-MgUser -UserId $UPN -PasswordProfile $Passwordprofile -ErrorAction Stop
  $log>>$PasswordLogFile
  $Script:ResetPasswordToRandomAction = "Success"
 }
 catch
 {
  $Script:ResetPasswordToRandomAction ="Failed"
  $ErrorLog = "$($UPN) - Reset Password To Random Action - "+$Error[0].Exception.Message
  $ErrorLog>>$ErrorsLogFile
 }
}

Function ReviewMFAMethods
{
 Try
 {
  $AuthenticationMethods = Get-MgUserAuthenticationMethod -UserId $UPN -ErrorAction Stop
  $MethodTypes = $AuthenticationMethods | ForEach-Object {
   $_.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.', ''
  } | Where-Object { $_ -ne 'passwordAuthenticationMethod' }
  If(($MethodTypes | Measure-Object).Count -eq 0)
  {
   $Script:ReviewMFAMethodsAction ="MFA is not registered. It's recommended to enable MFA"
  }
  else
  {
   $MethodTypes=$MethodTypes -join ","
   $Script:ReviewMFAMethodsAction = "Registered authentication methods: $MethodTypes"
  }
 }
 Catch
 {
  $Script:ReviewMFAMethodsAction ="Failed"
  $ErrorLog = "$($UPN) - Review registered MFA authentication method - "+$Error[0].Exception.Message
  $ErrorLog>>$ErrorsLogFile
 }
}

Function DisableInboxRule
{
    if($MailBoxAvailability -eq 'No')
    {
        $Script:DisableInboxRuleAction = "No Exchange license assigned to user"
        return
    }
    try{
        $MailboxRule = Get-InboxRule -Mailbox $UPN | where{$_.Enabled -eq $true}
        $RuleCount=($MailboxRule | Measure-Object).Count
        if($RuleCount -eq 0)
        {
         $Script:DisableInboxRuleAction = "No inbox rules or no rules are in enabled status"
        }
        else
        {
         $RuleNames=$MailboxRule.Name
         $RuleNames=$RuleNames -join ","
         $MailboxRule| Disable-inboxRule -Confirm:$False -ErrorAction Stop
         $Script:DisableInboxRuleAction = "Successfully disabled $RuleCount inbox rules: $RuleNames"
        }
    }
    catch
    {
     $Script:DisableInboxRuleAction = "Error occurred"
     $ErrorLog = "$($UPN) - Disable inbox rules - "+ $_.Exception.Message
     $ErrorLog>>$ErrorsLogFile
    }
}

Function ReviewEmailForwardingConfiguredByUser
{
 if($MailBoxAvailability -eq 'No')
 {
  $Script:ReviewEmailForwardingConfiguredByUserAction = "No Exchange license assigned to user"
  return
 }
 
 $ForwardingSMTPAddress=$MailBox.ForwardingSmtpAddress
 if($ForwardingSMTPAddress -eq $null)
 {
  $Script:ReviewEmailForwardingConfiguredByUserAction = "ForwardingSMTPAddress is not configured. You can verify inbox rules for external email forwarding configuration."
 }
 else
 {
  $ForwardingSMTPAddress=$ForwardingSMTPAddress.split(":") | Select -Index 1
  $Script:ReviewEmailForwardingConfiguredByUserAction = "$ForwardingSMTPAddress is configured to forward emails. You can verify inbox rules for additional email forwarding configuration"
 }
}

Function RemoveEmailForwardingConfiguredByUser
{
 if($MailBoxAvailability -eq 'No')
 {
  $Script:RemoveEmailForwardingConfiguredByUserAction = "No Exchange license assigned to user"
  return
 }
 try
 {
  Set-Mailbox -Identity $UPN -ForwardingSmtpAddress $null -ErrorAction Stop
  $Script:RemoveEmailForwardingConfiguredByUserAction="Success"
 }
 catch
 {
  $Script:RemoveEmailForwardingConfiguredByUserAction = "Failed"
  $ErrorLog = "$($UPN) - Remove Email Forwarding Configuration - "+$Error[0].Exception.Message
  $ErrorLog>>$ErrorsLogFile
 } 
}

Function GetActivityLog
{
  $EndDate=(Get-Date).AddSeconds(-1)
  $StartDate=(Get-date).AddDays(-10).Date
  $OutputCSV="$Location\$DisplayName _ActivityReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
  $AuditRecordCount=0
  $ErrorInfo=""
  Search-UnifiedAuditLog -UserIds $UPN -ResultSize 5000 -SessionCommand ReturnLargeSet -StartDate $StartDate -EndDate $EndDate | foreach {
   $AuditRecordCount++
   if($AuditRecordCount -eq 5000)
   {
    $ErrorInfo="More audit records available to fetch for $UPN. Reduce shorter interval to fetch complete data."
    Write-Host $ErrorInfo -ForegroundColor Red
   }
   $EventTime=$_.CreationDate
   $UserName=$_.UserIds
   $Operation=$_.Operations
   $AuditData=$_.AuditData 
   $AuditInfo=$AuditData| ConvertFrom-Json
   $Workload=$AuditInfo.Workload
   $ResultStatus=$AuditInfo.ResultStatus
   $Result=[PSCustomObject]@{'Activity Time'=$EventTime;'User Name'=$UserName;'Operation'=$Operation;'Result'=$ResultStatus;'Workload'=$Workload;'More Info'=$AuditData}
   $Result | Export-Csv -Path $OutputCSV -Notype -Append
  }
  if($AuditRecordCount -eq 0)
  {
   $Script:GetActivitytLogAction= "No activity found in the last 10 days"
  }
  else
  {
   $Script:GetActivitytLogAction = "Audit log available in $OutputCSV.$ErrorInfo"
  }
}
Function Disconnect_Modules
{
    Disconnect-MgGraph -ErrorAction SilentlyContinue|  Out-Null
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

Function main
{ 
   ConnectModules
    #Importing CSV file
    if($CSVFilePath -ne "")
    {
        $CSVFilePath = $CSVFilePath.Trim()
        try{
            $UPNCSVFile = Import-Csv -Path $CSVFilePath -Header UserPrincipalName
            [array]$UPNs = $UPNCSVFile.UserPrincipalName
        }
        catch
        {
            Write-Host $_.Exception.Message -ForegroundColor Red
            Exit
        }
    }
    elseif($UPNs -ne "")
    {
        [array]$UPNs = $UPNs.Split(',')
    }
    else
    {
        $UPNs = Read-Host `nEnter the UserPrincipalName of the compromised user
        if($UPNs -ne "")
        {
            [array]$UPNs = $UPNs -split ','
        }
        else
        {
            Write-Host You must provide UPN of the compromised user. -ForegroundColor Red
            Disconnect_Modules
        }
    }
    $script:Location = Get-Location
    $ExportCSV =  "$Location\CompromisedUser_remediation_StatusFile_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
    $PasswordLogFile = "$Location\PasswordLogFile_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).txt"
    $InvalidUserLogFile = "$Location\InvalidUsersLogFile$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).txt"
    $ErrorsLogFile = "$Location\ErrorsLogFile$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).txt"
    $AvailabiltyOfInvalidUser = $false
    
    Write-Host "`nChoose the actions to perform.`n" -ForegroundColor Cyan
    Write-Host "           1.  Disable user" -ForegroundColor Yellow
    Write-Host "           2.  Sign-out from all sessions" -ForegroundColor Yellow
    Write-Host "           3.  Reset password to random" -ForegroundColor Yellow 
    Write-Host "           4.  Review MFA registered method" -ForegroundColor Yellow
    Write-Host "           5.  Disable inbox rules" -ForegroundColor Yellow
    Write-Host "           6.  Review email forwarding configuration" -ForegroundColor Yellow
    Write-Host "           7.  Remove email forwarding" -ForegroundColor Yellow
    Write-Host "           8.  Get activity audit log for 10 days" -ForegroundColor Yellow
    Write-Host "           9. All the above operations" -ForegroundColor Yellow
    $Actions=Read-Host "`nPlease choose the action to continue"
    if($Actions -eq "")
    {
        Write-Host "`nPlease choose the action from the above." -ForegroundColor Red
        Exit
    }
    $Actions = $Actions.Trim()
    $Actions = $Actions.Split(',')
    $CheckActions = Compare-Object -Referenceobject $Actions -DifferenceObject @(1..9)
    if($CheckActions|?{$_.SideIndicator -eq "<="})
    {
        Write-Host "`nPlease choose the correct action number from the above actions." -ForegroundColor Red
        Disconnect_Modules
    }
    Foreach($UPN in $UPNs)
    {
        $UPN = $UPN.Trim()
        Write-Progress -Activity "Processing $UPN"
        $Script:Status = "$UPN - "
        $User = Get-MgUser -UserId $UPN -ErrorAction SilentlyContinue 
        $UserId = $User.Id
        $script:DisplayName=$User.DisplayName
        if($User -eq $null)
        {
            $InvalidUser= "$UPN"
            $InvalidUser>>$InvalidUserLogFile
            Continue
        }
        $MailBox = Get-Mailbox -Identity $UPN -RecipientTypeDetails UserMailbox -ErrorAction SilentlyContinue
        if($MailBox -ne $null)
        {
            $MailBoxAvailability = "Yes"
        }
        else
        {
            $MailBoxAvailability = "No"
        }
        if($Actions -contains 9)
        {
            $Actions = 1..8
        }
         foreach($Action in $Actions)
        {
            switch($Action){
                1  {   
                      Write-Progress -Activity "Processing $UPN" -Status "Disabling user"  
                      DisableUser ; break 
                   }
                2  {   
                      Write-Progress -Activity "Processing $UPN" -Status "Signing out from all the sessions" 
                      SignOutFromAllSessions ; break 
                   }
                3  {  
                      Write-Progress -Activity "Processing $UPN" -Status "Resetting password" 
                      ResetPasswordToRandom ; break 
                   }
                4  {   
                      Write-Progress -Activity "Processing $UPN" -Status "Retrieving registered authentication methods"
                      ReviewMFAMethods ; break 
                   }
                5  {   
                      Write-Progress -Activity "Processing $UPN" -Status "Disable inbox rules"
                      DisableInboxRule ; break 
                   }
                6  {  
                      Write-Progress -Activity "Processing $UPN" -Status "Retrieving SMTP email forwarding configuration set by user"
                      ReviewEmailForwardingConfiguredByUser ; break 
                   }
                7  {  
                      Write-Progress -Activity "Processing $UPN" -Status "Removing SMTP forwarding address"
                      RemoveEmailForwardingConfiguredByUser ; break 
                   }
                8  {  
                      Write-Progress -Activity "Processing $UPN" -Status "Exporting activity audit logs"
                      GetActivityLog ; break 
                   }
               
                Default {
                    Write-Host "No action found. Please provide valid input" -ForegroundColor Red
                    Disconnect_Modules
                }
            }
        }
        #This is for to set mailbox availablity value only if actions contains mailbox related actions. Otherwise its value set to be null.
        $MailboxRelatedActions = Compare-Object $Actions @(5,6,7) -IncludeEqual -ExcludeDifferent
        if($MailboxRelatedActions.count -eq 0)
        {
            $MailBoxAvailability = ""
        }
        if($MailBoxAvailability -eq "No")
        {
            $MailBoxAvailability = "No Exchange license assigned to user"
        }
        $Result = [PSCustomObject]@{
        'UPN'=$UPN;'Disable User'=$DisableUserAction;'Reset Password To Random'=$ResetPasswordToRandomAction;'Exchange User'=$MailBoxAvailability;'Disable Inbox Rule'=$DisableInboxRuleAction;'SignOut From All Sessions'=$SignOutFromAllSessionsAction;'Registered MFA Methods'=$ReviewMFAMethodsAction;'Review Email Forwarding Set by User'=$ReviewEmailForwardingConfiguredByUserAction;'Remove Email Forwarding Set by User'=$RemoveEmailForwardingConfiguredByUserAction;'Export Audit Log'=$GetActivitytLogAction} 
        $Result | Export-csv -Path $ExportCSV -Append -NoTypeInformation
        $Variables = @("DisableUserAction","SignOutFromAllSessionsAction","ResetPasswordToRandomAction","MailBoxAvailability","DisableInboxRuleAction","ReviewMFAMethodsAction","ReviewEmailForwardingConfiguredByUserAction","RemoveEmailForwardingConfiguredByUserAction","GetActivitytLogAction")
        $Variables | ForEach-Object {Clear-Variable -Name $_ -ErrorAction SilentlyContinue}
    }

    Write-Host `nScript execution completed 

    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 
    if(Test-Path -Path $ExportCSV)
    {
        Write-Host `nStatus file available in -NoNewline -Foregroundcolor Yellow; Write-Host " $ExportCSV" 
        $Prompt = New-Object -ComObject wscript.shell  
        $UserInput = $Prompt.popup("Do you want to open output file?",`  0,"Open Output File",4)  
        if ($UserInput -eq 6)  
        {  
            Invoke-Item "$ExportCSV"  
        } 
    }
    if(Test-Path -Path $InvalidUserLogFile)
    {
        Write-Host `nInvalid users log file available in -NoNewline -Foregroundcolor Yellow; Write-Host " $InvalidUserLogFile" 
    }
    if(Test-Path -Path $ErrorsLogFile)
    {
        Write-Host `nErrors log file available in -NoNewline -Foregroundcolor Yellow; Write-Host " $InvalidUserLogFile .You can see the reason for failed actions." 
    }
    if(Test-Path -Path $PasswordLogFile)
    {
        Write-Host `nPassword log file available in -NoNewline -Foregroundcolor Yellow; Write-Host " $PasswordLogFile" 
    }
   ############################################   Disconnect_Modules
}


. main