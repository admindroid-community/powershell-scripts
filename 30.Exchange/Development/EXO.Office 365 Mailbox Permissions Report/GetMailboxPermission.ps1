<#
=============================================================================================
Name:           Export Mailbox Permission Report
Website:        o365reports.com
Version:        3.0

Script Highlights :
~~~~~~~~~~~~~~~~~

1. The script uses Modern authentication to connect to Exchange Online.
2. The script display only “Explicitly assigned permissions” to mailboxes which means it will ignore “SELF” permission that each user on his mailbox and inherited permission.
3. Exports output to CSV file.
4. The script can be executed with MFA enabled account too.
5. The script supports certificate based authentication (CBA) too.
6. You can choose to either “export permissions of all mailboxes” or pass an input file to get permissions of specific mailboxes alone.
7. Allows you to filter output using your desired permissions like Send-as, Send-on-behalf or Full access.
8. Output can be filtered based on user/all mailbox type
9. Allows you to filter permissions on admin’s mailbox. So that you can view administrative users’ mailbox permission alone.
10. Automatically installs the EXO V2 and MS Graph PowerShell modules (if not installed already) upon your confirmation. 
11. This script is scheduler friendly.


For detailed Script execution: https://o365reports.com/2019/03/07/export-mailbox-permission-csv/
============================================================================================
#>

#If you connect via Certificate based authentication, then your application required "Directory.Read.All" application permission, assign exchange administrator role and  Exchange.ManageAsApp permission to your application.

param(
[switch]$FullAccess,
[switch]$SendAs,
[switch]$SendOnBehalf,
[switch]$UserMailboxOnly,
[switch]$AdminsOnly,
[string]$MBNamesFile,
[string]$TenantId,
[string]$ClientId,
[string]$CertificateThumbprint
)

Function ConnectModules 
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
    Write-Progress -Activity "Connecting modules(Microsoft Graph and Exchange Online module)..."
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
            if($Scopes -notcontains "Directory.Read.All" -and $Scopes -notcontains "Directory.ReadWrite.All")
            {
                Write-Host "Note: Your application required the following graph application permissions: Directory.Read.All" -ForegroundColor Yellow
                Exit
            }
            Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization (Get-MgDomain | Where-Object {$_.isInitial}).Id -ShowBanner:$false
        }
        else
        {
            Connect-MgGraph -Scopes "Directory.Read.All"  -ErrorAction SilentlyContinue -Errorvariable ConnectionError |Out-Null
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
    Write-Host "Microsoft Graph Beta PowerShell module is connected successfully" -ForegroundColor Cyan
    Write-Host "Exchange Online module is connected successfully" -ForegroundColor Cyan
}
Function Print_Output
{
    $Result = [PSCustomObject]@{'DisplayName'=$Displayname;'UserPrincipalName'=$UPN;'MailboxType'=$MBType;'AccessType'=$AccessType;'UserWithAccess'=$UserWithAccess;'Roles'=$Roles} 
    $Result | Export-csv -Path $ExportCSV -Append -NoTypeInformation
}
Function Get_MBPermission
{
    #Getting delegated Fullaccess permission for mailbox
    if(($FilterPresent -eq 'False') -or ($FullAccess.IsPresent))
    {
        $FullAccessPermissions = (Get-EXOMailboxPermission -Identity $UPN -ErrorAction SilentlyContinue | Where { ($_.AccessRights -contains "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -match "NT AUTHORITY" -or $_.User -match "S-1-5-21") }).User
        if([string]$FullAccessPermissions -ne "")
        {
            $AccessType = "FullAccess"
            $UserWithAccess = @($FullAccessPermissions) -join ','
            Print_Output
        }
    }
    #Getting delegated SendAs permission for mailbox
    if(($FilterPresent -eq 'False') -or ($SendAs.IsPresent))
    {
        $SendAsPermissions = (Get-EXORecipientPermission -Identity $UPN -ErrorAction SilentlyContinue | Where{ -not (($_.Trustee -match "NT AUTHORITY") -or ($_.Trustee -match "S-1-5-21"))}).Trustee
        if([string]$SendAsPermissions -ne "")
        {
            $AccessType = "SendAs"
            $UserWithAccess = @($SendAsPermissions) -join ','
            Print_Output
        }
    }
    #Getting delegated SendOnBehalf permission for mailbox
    if(($FilterPresent -eq 'False') -or ($SendOnBehalf.IsPresent))
    {
        if([string]$SendOnBehalfPermissions -ne "")
        {
            $AccessType = "SendOnBehalf"
            $UserWithAccess = @()
            Foreach($SendOnBehalfPermissionDN in $SendOnBehalfPermissions)
            {
                $SendOnBehalfPermission = (Get-EXOMailBox -Identity $SendOnBehalfPermissionDN -ErrorAction SilentlyContinue).UserPrincipalName
                if($SendOnBehalfPermission -eq $null)
                {
                    $SendOnBehalfPermission = ($Users|?{$_.MailNickname -eq $SendOnBehalfPermissionDN}).UserPrincipalName
                }
                $UserWithAccess += $SendOnBehalfPermission
            }
            $UserWithAccess = @($UserWithAccess) -join ','
            Print_Output
        }
    }
}
#Getting Mailbox permission
Function Get_MailBoxData
{
    Write-Progress -Activity "`n     Processing mailbox: $MBUserCount `n  Currently Processing: $DisplayName" 
    $Script:MBUserCount++
    if($UserMailboxOnly.IsPresent -and $MBType -ne 'UserMailBox')
    {
        return
    }
    #Get admin roles assigned to user 
    $RoleList=Get-MgBetaUserTransitiveMemberOf -UserId $UPN|Select-Object -ExpandProperty AdditionalProperties
    $RoleList = $RoleList|?{$_.'@odata.type' -eq '#microsoft.graph.directoryRole'}
    $Roles = @($RoleList.displayName) -join ','
    if($RoleList.count -eq 0)
    {
        $Roles = "No roles"
    }
    
    #Admin Role based filter
    if($AdminsOnly.IsPresent -and $Roles -eq "No roles")
    { 
        return
    }
    Get_MBPermission
}
Function CloseConnection
{
    Disconnect-MgGraph | Out-Null 
    Disconnect-ExchangeOnline -Confirm:$false
}
ConnectModules
Write-Host "`nNote: If you encounter module related conflicts, run the script in a fresh PowerShell window." -ForegroundColor Yellow

Write-Progress -Activity Completed -Completed
#Set output file
$Location = (Get-Location)
$ExportCSV =  "$($Location)\MBPermission_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$Result = "" ; $Mailboxes = @(); $MBUserCount = 1;
$Users = Get-MgBetaUser -All
#Check for AccessType filter
if(($FullAccess.IsPresent) -or ($SendAs.IsPresent) -or ($SendOnBehalf.IsPresent))
{
    $FilterPresent = 'True'
}
else
{
    $FilterPresent = 'False'
}

#Check for input file
if ($MBNamesFile -ne "") 
{ 
    #We have an input file, read it into memory 
    try{
        $MailBoxes = Import-Csv -Header "MailBoxUPN" -Path $MBNamesFile
    }
    catch{
        Write-Host $_.Exception.Message -ForegroundColor Red
        CloseConnection
        Exit
    }
    Foreach($Mail in $MailBoxes)
    {
        $Mailbox = Get-EXOMailbox -Identity $Mail.MailBoxUPN -PropertySets All -ErrorAction SilentlyContinue
        if($Mailbox -eq $null)
        {
            Write-Host `n $Mail.MailBoxUPN is not found -ForegroundColor Red
            Continue
        }
        $DisplayName = $MailBox.DisplayName
        $UPN = $MailBox.UserPrincipalName
        $MBType = $MailBox.RecipientTypeDetails
        $SendOnBehalfPermissions = $MailBox.GrantSendOnBehalfTo
        Get_MailBoxData
    }
}
else
{
    Get-EXOMailbox -ResultSize Unlimited -PropertySets All | Where{$_.DisplayName -notlike "Discovery Search Mailbox"} |ForEach-Object{
        $DisplayName = $_.DisplayName
        $UPN = $_.UserPrincipalName
        $MBType = $_.RecipientTypeDetails
        $SendOnBehalfPermissions = $_.GrantSendOnBehalfTo   
        Get_MailBoxData
    }
}
#Open output file after execution 
Write-Host `nScript executed successfully
if((Test-Path -Path $ExportCSV) -eq "True")
{
    Write-Host Detailed report available in: -NoNewline -Foregroundcolor Yellow; Write-Host " $ExportCSV" 
    $Prompt = New-Object -ComObject wscript.shell  
    $UserInput = $Prompt.popup("Do you want to open output file?",`  0,"Open Output File",4)  
    if ($UserInput -eq 6)  
    {  
        Invoke-Item "$ExportCSV"  
    } 
}
else
{
    Write-Host No mailbox found that matches your criteria. -ForegroundColor Red
}
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
CloseConnection