<#
=============================================================================================
Name:           Automate Microsoft 365 User Offboarding with PowerShell
Description:    This script can perform 14 Microsoft 365 offboarding activities.
Website:        blog.Admindroid.com
Script by:      AdminDroid Team
Version:        2.0


For detailed Script execution: https://blog.admindroid.com/automate-microsoft-365-user-offboarding-with-powershell


Change Log
~~~~~~~~~~

    V1.0 (Oct 14, 2023) - File created
    V2.0 (Apr 02, 2025) - Removed beta version cmdlets 

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
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Disconnect-ExchangeOnline -Confirm:$false
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
            $Scopes = (Get-MgContext).Scopes
            $ApplicationPermissions=@("Directory.ReadWrite.All","AppRoleAssignment.ReadWrite.All","User.EnableDisableAccount.All","RoleManagement.ReadWrite.Directory")
            foreach($Permission in $ApplicationPermissions)
            {
                if($Scopes -notcontains $Permission)
                {
                    Write-Host "Note: Your application required the following graph application permissions: Directory.ReadWrite.All,AppRoleAssignment.ReadWrite.All,User.EnableDisableAccount.All,RoleManagement.ReadWrite.Directory" -ForegroundColor Yellow
                    Exit
                }
            }
            Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization (Get-MgDomain | Where-Object {$_.isInitial}).Id -ShowBanner:$false
        }
        else
        {
            Connect-MgGraph -Scopes Directory.ReadWrite.All,AppRoleAssignment.ReadWrite.All,User.EnableDisableAccount.All,Directory.AccessAsUser.All,RoleManagement.ReadWrite.Directory -ErrorAction SilentlyContinue -Errorvariable ConnectionError |Out-Null
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
        Update-MgUser -UserId $UPN -AccountEnabled:$false
        $Script:DisableUserAction = "Success"
    }
    catch
    {
        $Script:DisableUserAction = "Failed"
        $ErrorLog = "$($UPN) - Disable User Action - "+$Error[0].Exception.Message
        $ErrorLog>>$ErrorsLogFile
    }
}

Function ResetPasswordToRandom
{
    $Password = -join ((48..57) + (65..90) + (97..122) | ForEach-Object { [char]$_ } | Get-Random -Count 8)
    $log = "$UPN - $Password"
    $Pwd = ConvertTo-SecureString $Password -AsPlainText –Force
    try{
        $Passwordprofile = @{
		    forceChangePasswordNextSignIn = $true
		    password = $Pwd
	    }
        Update-MgUser -UserId $UPN -PasswordProfile $Passwordprofile
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
Function ResetOfficeName
{
    try{
        Update-MgUser -UserId $UPN -OfficeLocation "EXD"
        $Script:ResetOfficeNameAction = "Success"
    }
    catch
    {
        $Script:ResetOfficeNameAction = "Failed"
        $ErrorLog = "$($UPN) - Reset Office Name Action - "+$Error[0].Exception.Message
        $ErrorLog>>$ErrorsLogFile
    }
}

Function RemoveMobileNumber
{
    try{
        Update-MgUser -UserId $UPN -MobilePhone null
        $Script:RemoveMobileNumberAction = "Success"
    }
    catch
    {
        $Script:RemoveMobileNumberAction = "Failed"
        $ErrorLog = "$($UPN) - Remove Mobile Number Action - "+$Error[0].Exception.Message
        $ErrorLog>>$ErrorsLogFile
    }
}

Function RemoveGroupMemberships
{
    #Remove memberships from group
    $groupMemberships = $Memberships|?{($_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') -and ($_.AdditionalProperties.'groupTypes' -notcontains 'DynamicMembership')}
    foreach($Membership in $groupMemberships)
    {
        try{ 
            Remove-MgGroupMemberByRef -GroupId $Membership.Id -DirectoryObjectId $UserId -ErrorAction SilentlyContinue -ErrorVariable MemberRemovalErr
            if($MemberRemovalErr)
            {
                Remove-DistributionGroupMember -Identity $Membership.Id  -Member $UserId -BypassSecurityGroupManagerCheck -Confirm:$false
            }
        }
        catch
        {
            $ErrorLog = "$($UPN) - GroupId($($Membership.Id)) - Remove Group Memberships Action - "+$Error[0].Exception.Message
            $ErrorLog>>$ErrorsLogFile
        }
    }
    #Remove ownerships from group
    $GroupOwnerships = Get-MgUserOwnedObject -UserId $UPN|?{$_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'}
    foreach($GroupOwnership in $GroupOwnerships)
    {
        try{
            Remove-MgGroupOwnerByRef -GroupId $GroupOwnership.Id -DirectoryObjectId $UserId -ErrorAction SilentlyContinue -ErrorVariable OwnerRemovalErr
            if($OwnerRemovalErr)
            {
                $ErrorLog = "$($UPN) - GroupId($($GroupOwnership.Id)) - Remove Group Memberships Action - "+$OwnerRemovalErr.Exception.Message
                $ErrorLog>>$ErrorsLogFile
            }
        }
        catch
        {
            $ErrorLog = "$($UPN) - GroupId($($GroupOwnership.Id)) - Remove Group Memberships Action - "+$Error[0].Exception.Message
            $ErrorLog>>$ErrorsLogFile
        }
    }
    $DistributionGroupOwnerships = Get-DistributionGroup | where {$_.ManagedBy -contains "$UserId"}
    foreach($DistributionGroupOwnership in $DistributionGroupOwnerships)
    {
        Set-DistributionGroup -Identity $DistributionGroupOwnership.Identity -BypassSecurityGroupManagerCheck -ManagedBy @{Remove=$UPN} -ErrorAction SilentlyContinue -ErrorVariable OwnerRemovalErr
        if($OwnerRemovalErr)
        {
            $ErrorLog = "$($UPN) - GroupId($($DistributionGroupOwnership.ExternalDirectoryObjectId)) - Remove Group Memberships Action - "+$OwnerRemovalErr.Exception.Message
            $ErrorLog>>$ErrorsLogFile
        }
    }
    if($ErrorLog -eq $null)
    {
        $Script:RemoveGroupMembershipsAction = "Success"
    }
    elseif($groupMemberships -eq $null -and $GroupOwnerships -eq $null -and $DistributionGroupOwnerships -eq $null)
    {
        $Script:RemoveGroupMembershipsAction = "No group memberships"
    }
    else
    {
        $Script:RemoveGroupMembershipsAction = "Failed"
    }

}

Function RemoveAdminRoles
{
    $AdminRoles = $Memberships|?{$_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.directoryRole'}
    if($AdminRoles -eq $null)
    {
        $Script:RemoveAdminRolesAction = "No admin roles"
    }
    else
    {
        foreach($AdminRole in $AdminRoles)
        {
            try{
                Remove-MgDirectoryRoleMemberByRef -DirectoryObjectId $UserId -DirectoryRoleId $AdminRole.Id 
            }
            catch
            {
                $ErrorLog = "$($UPN) - Role Id($($Role.DisplayName)) Remove Admin Roles Action - "+$Error[0].Exception.Message
                $ErrorLog>>$ErrorsLogFile
            }
        }
        if($ErrorLog -eq $null)
        {
            $Script:RemoveAdminRolesAction = "Success"
        }
        else
        {
            $Script:RemoveAdminRolesAction = "Failed"
        }
    }
}
Function RemoveAppRoleAssignments
{
    $AppRoleAssignments = Get-MgUserAppRoleAssignment -UserId $UPN
    if($AppRoleAssignments -ne $null)
    {
        $AppRoleAssignments | ForEach-Object {
            try{
                Remove-MgUserAppRoleAssignment -AppRoleAssignmentID $_.Id -UserId $UPN
            }
            catch
            {
                $ErrorLog = "$($UPN) - Remove App Role Assignments Action - "+$Error[0].Exception.Message
                $ErrorLog>>$ErrorsLogFile
            }
        }
        if($ErrorLog -eq $null)
        {
            $Script:RemoveAppRoleAssignmentsAction = "Success"
        }
        else
        {
            $Script:RemoveAppRoleAssignmentsAction = "Failed"
        }
    }
    else
    {
        $Script:RemoveAppRoleAssignmentsAction = "No app role assignments"
    }
}

Function HideFromAddressList
{
    if($MailBoxAvailability -eq 'No')
    {
        $Script:HideFromAddressListAction = "No Exchange license assigned to user"
        return
    }
    try{
        Set-Mailbox -Identity $UPN -HiddenFromAddressListsEnabled $true 
        $Script:HideFromAddressListAction = "Success"
    }
    catch
    {
        $Script:HideFromAddressListAction = "Failed"
        $ErrorLog = "$($UPN) - Hide From Address List Action - "+$Error[0].Exception.Message
        $ErrorLog>>$ErrorsLogFile
    }
}
Function RemoveEmailAlias
{
    if($MailBoxAvailability -eq 'No')
    {
        $Script:RemoveEmailAliasAction = "No Exchange license assigned to user"
        return
    }
    try{
        $EmailAliases=Get-Mailbox $UPN| select -ExpandProperty emailaddresses| ?{$_.StartsWith("smtp")}
        if($EmailAliases -eq $null)
        {
            $Script:RemoveEmailAliasAction = "No alias"
        }
        else
        {
            Set-Mailbox $UPN -EmailAddresses @{Remove=$EmailAliases} -WarningAction SilentlyContinue
            $Script:RemoveEmailAliasAction = "Success"
        }
    }
    catch
    {
        $Script:RemoveEmailAliasAction = "Failed"
        $ErrorLog = "$($UPN) - Remove Email Alias Action - "+$Error[0].Exception.Message
        $ErrorLog>>$ErrorsLogFile
    }
}

Function WipingMobileDevice
{
    if($MailBoxAvailability -eq 'No')
    {
        $MobileDeviceAction = "No Exchange license assigned to user"
        return
    }
    try{
        $MobileDevice = Get-MobileDevice -Mailbox $UPN 
        $MobileDevice| Clear-MobileDevice
        $Script:MobileDeviceAction = "Success"
    }
    catch
    {
        $Script:MobileDeviceAction = "Failed"
        $ErrorLog = "$($UPN) - Wiping Mobile Device Action - "+$Error[0].Exception.Message
        $ErrorLog>>$ErrorsLogFile
    }
}

Function DeleteInboxRule
{
    if($MailBoxAvailability -eq 'No')
    {
        $Script:DeleteInboxRuleAction = "No Exchange license assigned to user"
        return
    }
    try{
        $MailboxRule = Get-InboxRule -Mailbox $UPN 
        $MailboxRule| Remove-InboxRule -Confirm:$False
        $Script:DeleteInboxRuleAction = "Success"
    }
    catch
    {
        $Script:DeleteInboxRuleAction = "No inbox rule"
    }
}

Function ConvertToSharedMailbox
{
    if($MailBoxAvailability -eq 'No')
    {
        $Script:ConvertToSharedMailboxAction = "No Exchange license assigned to user"
        return
    }
    try{
        Set-Mailbox -Identity $UPN -Type Shared -WarningAction SilentlyContinue
        $Script:ConvertToSharedMailboxAction = "Success"
    }
    catch
    {
        $Script:ConvertToSharedMailboxAction = "Failed"
        $ErrorLog = "$($UPN) - Convert To Shared Mailbox Action - "+$Error[0].Exception.Message
        $ErrorLog>>$ErrorsLogFile
    }
}

Function RemoveLicense
{
    $Licenses = Get-MgUserLicenseDetail -UserId $UPN
    if($Licenses -ne $null)
    {
        Set-MgUserLicense -UserId $UPN -RemoveLicenses @($Licenses.SkuId) -AddLicenses @() -ErrorAction SilentlyContinue -ErrorVariable LicenseError | Out-Null
        if($LicenseError)
        {
            $Script:RemoveLicenseAction = "Failed"
            $ErrorLog = "$($UPN) - Remove License Action - "+$LicenseError.Exception.Message 
            $ErrorLog>>$ErrorsLogFile
        }
        else
        {
            $Script:RemoveLicenseAction = "Removed licenses - $($Licenses.SkuPartNumber -join ',')"
        }
    }
    else
    {
        $Script:RemoveLicenseAction = "No license"
    }
}

Function SignOutFromAllSessions
{
    Revoke-MgUserSignInSession -UserId $UPN | Out-Null
    $Script:SignOutFromAllSessionsAction = "Success"
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
        $UPNs = Read-Host `nEnter the UserPrincipalName of the user you want to offboard
        if($UPNs -ne "")
        {
            [array]$UPNs = $UPNs -split ','
        }
        else
        {
            Write-Host You must provide UPN of the user to offboard. -ForegroundColor Red
            Disconnect_Modules
        }
    }
    $Location = Get-Location
    $ExportCSV =  "$Location\M365UserOffBoarding_StatusFile_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
    $PasswordLogFile = "$Location\PasswordLogFile_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).txt"
    $InvalidUserLogFile = "$Location\InvalidUsersLogFile$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).txt"
    $ErrorsLogFile = "$Location\ErrorsLogFile$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).txt"
    $AvailabiltyOfInvalidUser = $false
    
    Write-Host "`nWe can perform below operations.`n" -ForegroundColor Cyan
    Write-Host "           1.  Disable user" -ForegroundColor Yellow
    Write-Host "           2.  Reset password to random" -ForegroundColor Yellow 
    Write-Host "           3.  Reset Office name" -ForegroundColor Yellow 
    Write-Host "           4.  Remove Mobile number" -ForegroundColor Yellow
    Write-Host "           5.  Remove group memberships" -ForegroundColor Yellow
    Write-Host "           6.  Remove admin roles" -ForegroundColor Yellow
    Write-Host "           7.  Remove app role assignments" -ForegroundColor Yellow
    Write-Host "           8.  Hide from address list" -ForegroundColor Yellow
    Write-Host "           9.  Remove email alias" -ForegroundColor Yellow
    Write-Host "           10. Wiping mobile device" -ForegroundColor Yellow
    Write-Host "           11. Delete inbox rule" -ForegroundColor Yellow
    Write-Host "           12. Convert to shared mailbox" -ForegroundColor Yellow
    Write-Host "           13. Remove license" -ForegroundColor Yellow
    Write-Host "           14. Sign-out from all sessions" -ForegroundColor Yellow
    Write-Host "           15. All the above operations" -ForegroundColor Yellow
    $Actions=Read-Host "`nPlease choose the action to continue"
    if($Actions -eq "")
    {
        Write-Host "`nPlease choose the action from the above." -ForegroundColor Red
        Exit
    }
    $Actions = $Actions.Trim()
    $Actions = $Actions.Split(',')
    $CheckActions = Compare-Object -Referenceobject $Actions -DifferenceObject @(1..15)
    if($CheckActions|?{$_.SideIndicator -eq "<="})
    {
        Write-Host "`nPlease choose the correct action number from the above actions." -ForegroundColor Red
        Disconnect_Modules
    }
    Foreach($UPN in $UPNs)
    {
        $UPN = $UPN.Trim()
        Write-Progress "Processing $UPN"
        $Script:Status = "$UPN - "
        $User = Get-MgUser -UserId $UPN -ErrorAction SilentlyContinue 
        $UserId = $User.Id
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
        if($Actions -contains 15)
        {
            $Actions = 1..14
        }
        if($Actions -contains 5 -or $Actions -contains 6) # To get memberships of the user (group and roles)
        {
            $Memberships = Get-MgUserMemberOf -UserId $UPN
        }       
        foreach($Action in $Actions)
        {
            switch($Action){
                1  {  DisableUser ; break }
                2  {  ResetPasswordToRandom ; break }
                3  {  ResetOfficeName ; break }
                4  {  RemoveMobileNumber ; break }
                5  {  RemoveGroupMemberships ; break }
                6  {  RemoveAdminRoles ; break }
                7  {  RemoveAppRoleAssignments ; break }
                8  {  HideFromAddressList ; break }
                9  {  RemoveEmailAlias ; break }
                10 {  WipingMobileDevice ; break }
                11 {  DeleteInboxRule ; break }
                12 {  ConvertToSharedMailbox ; break }
                13 {  RemoveLicense ; break }
                14 {  SignOutFromAllSessions ; break }
                Default {
                    Write-Host "No action found. Please provide valid input" -ForegroundColor Red
                    Disconnect_Modules
                }
            }
        }
        #This is for to set mailbox availablity value only if actions contains mailbox related actions. Otherwise its value set to be null.
        $MailboxRelatedActions = Compare-Object $Actions @(8,9,10,11,12,13) -IncludeEqual -ExcludeDifferent
        if($MailboxRelatedActions.count -eq 0)
        {
            $MailBoxAvailability = ""
        }
        if($MailBoxAvailability -eq "No")
        {
            $MailBoxAvailability = "No Exchange license assigned to user"
        }
        $Result = [PSCustomObject]@{
            'UPN'=$UPN;'Disable User'=$DisableUserAction;'Reset Password To Random'=$ResetPasswordToRandomAction;'Reset OfficeName'=$ResetOfficeNameAction;'Remove Mobile Number'=$RemoveMobileNumberAction;
            'Remove Group Memberships'=$RemoveGroupMembershipsAction;'Remove Admin Roles'=$RemoveAdminRolesAction;'Remove AppRole Assignments'=$RemoveAppRoleAssignmentsAction;'Exchange User'=$MailBoxAvailability;
            'Hide From Address List'= $HideFromAddressListAction;'Remove Email Alias'=$RemoveEmailAliasAction;'Wiping Mobile Device'=$MobileDeviceAction;
            'Delete Inbox Rule'=$DeleteInboxRuleAction;'ConvertToSharedMailbox'=$ConvertToSharedMailboxAction;'Remove License' =$RemoveLicenseAction;'SignOut From All Sessions'=$SignOutFromAllSessionsAction;} 
        $Result | Export-csv -Path $ExportCSV -Append -NoTypeInformation
        $Variables = @("DisableUserAction","ResetPasswordToRandomAction","ResetOfficeNameAction","RemoveMobileNumberAction","RemoveGroupMembershipsAction","RemoveAdminRolesAction","RemoveAppRoleAssignmentsAction","MailBoxAvailability","HideFromAddressListAction","RemoveEmailAliasAction","MobileDeviceAction","DeleteInboxRuleAction","ConvertToSharedMailboxAction","RemoveLicenseAction","SignOutFromAllSessionsAction")
        $Variables | ForEach-Object {Clear-Variable -Name $_ -ErrorAction SilentlyContinue}
    }
    Write-Host `nScript executed successfully -ForegroundColor Green
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
    if($Actions -contains 15 -or $Actions -contains 2)
    {
        Write-Host `nPassword log file available in -NoNewline -Foregroundcolor Yellow; Write-Host " $PasswordLogFile" 
    }
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    Disconnect_Modules
}
. main