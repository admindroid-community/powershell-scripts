<#
=============================================================================================
Name:           Upgrade Distribution Lists to Microsoft 365 Groups   Using PowerShell
Description:    This script converts distribution lists to Microsoft 365 groups
Version:        1.0
Website:        o365reports.com
Script by:      O365Reports Team
For detailed script execution: https://o365reports.com/2023/02/21/upgrade-distribution-lists-to-microsoft-365-groups-using-powershell/
============================================================================================
#>
param ( 
[string] $DistributionEmailAddress = $null,
[string] $UserName = $null, 
[string] $Password = $null,
[string] $InputFile = $null
) 
#Connect Modules
function Connect_ExchangeOnline
{
    Write-Progress -Activity "Connecting exchange online.."
    $ExchangeOnlineModule =  Get-Module ExchangeOnlineManagement -ListAvailable
    if($ExchangeOnlineModule -eq $null)
    { 
        Write-host "Important: Exchange online module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install ExchangeOnline module? [Y] Yes [N] No  
        if($confirm -match "[yY]") 
        { 
            Write-host "Installing exchange online module..."
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
            Write-host "Exchange online Module is installed in the machine successfully" -ForegroundColor Magenta 
        } 
        else
        { 
            Write-host "Exiting. `nNote: Exchange online module must be available in your system to run the script" 
            Exit 
        } 
    }
    Disconnect-ExchangeOnline -Confirm:$false
    try
    {
        if(($UserName -ne "") -and ($Password -ne ""))   
        {   
            $securedPassword = ConvertTo-SecureString -AsPlainText $Password -Force   
            $credential  = New-Object System.Management.Automation.PSCredential $UserName,$securedPassword   
            Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ErrorAction SilentlyContinue  # For Non-MFA account
        } 
        else
        {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction SilentlyContinue 
        }
    }
    catch
    {
        Write-Host $_.Exception.message -ForegroundColor Red
        Exit
    }
    Write-Host "Exchange online connected" -ForegroundColor Green
}
#To create group 
function Create_M365Group
{
    Write-Progress -Activity "Creating Microsoft 365 group...."
    $Params=@{}
    $Params= @{DisplayName                        =$DLGroup.DisplayName
               AccessType                         =$AccessType
               ManagedBy                          =$DLGroup.managedby  
               RequireSenderAuthenticationEnabled =$DLGroup.RequireSenderAuthenticationEnabled     
              }
    if($GroupMember -ne $null)
    {
        $Params.Add('Members',$GroupMember)
    }
    if($AccessType -eq 'Private')
    {
        $Params.Add('HiddenGroupMembershipEnabled',$DLGroup.HiddenGroupMembershipEnabled)
    }
    $script:NewM365Group = New-UnifiedGroup @Params
    if($NewM365Group -eq $null)
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        return
    }
    Write-Host "$($DLGroup.DisplayName) group created successfuly" -ForegroundColor Green
    $GetOwner = Compare-Object -ReferenceObject $DLGroup.managedby -DifferenceObject $NewM365Group.ManagedBy
    $GetOwner = $GetOwner |Where-Object{$_.SideIndicator -eq "=>"}
    $M365GroupOwner = $GetOwner.InputObject
    Remove-DistributionGroup -Identity $DLMail -Confirm:$false
    if($M365GroupOwner -ne $null)
    {
        Remove-UnifiedGroupLinks -Identity $NewM365Group.PrimarySmtpAddress -LinkType Owners -Links $M365GroupOwner -Confirm:$false -ErrorAction SilentlyContinue
        if($GroupMember -notcontains $M365GroupOwner)
        {
            Remove-UnifiedGroupLinks -Identity $NewM365Group.PrimarySmtpAddress -LinkType Members -Links $M365GroupOwner -Confirm:$false -ErrorAction SilentlyContinue
        }
    }
    While(1)
    {
        Write-Progress -Activity "Updating DL properties to Microsoft 365 group info..."
        Start-Sleep -Seconds 5
        Set-UnifiedGroup -Identity $NewM365Group.PrimarySmtpAddress -PrimarySmtpAddress $DLMail -ErrorAction SilentlyContinue -ErrorVariable DLGroupError
        if($DLGroupError -ne $null)
        {
            Write-Host "Removed existing Distribution List does not updated... Waiting for 5 seconds" -ForegroundColor Red
            Continue
        }
        Set-UnifiedGroup -Identity $DLMail -EmailAddresses @{Add = "X500:$($DLGroup.LegacyExchangeDN)"} -HiddenFromAddressListsEnabled $DLGroup.HiddenFromAddressListsEnabled -AcceptMessagesOnlyFromSendersOrMembers $DLGroup.AcceptMessagesOnlyFromSendersOrMembers -GrantSendOnBehalfTo $DLGroup.GrantSendOnBehalfTo -ModeratedBy $DLGroup.ModeratedBy -MailTip $DLGroup.MailTip  -ErrorAction SilentlyContinue
        break
    }
    Write-Host "$($DLGroup.DisplayName) group successfully converted to Microsoft 365 group" -ForegroundColor Green
    Update_NewM365GroupInfo
}
#Update the group info
function Update_NewM365GroupInfo
{ 
    $confirm = Read-Host Are you sure you want to update the Microsoft 365 group info? [Y] Yes [N] No
    if($confirm -match "[yY]") 
    { 
        Write-Progress -Activity "Updating Microsoft 365 group info..."
        $GroupName = Read-Host "Enter the Microsoft 365 group name"
        if($GroupName -ne "")
        {
            Set-UnifiedGroup -Identity $DLMail -DisplayName "$GroupName" -ErrorAction SilentlyContinue -ErrorVariable NameError
            if($NameError -ne $null)
            {
                Write-Host $NameError.Exception[1] -ForegroundColor Yellow
            }
            else
            {
                Write-Host "Group name updated successfully" -ForegroundColor Green
            }
        }
        $GroupMailAddress = Read-Host "Enter the Microsoft 365 group mail address"
        if($GroupMailAddress -ne "")
        {
            Set-UnifiedGroup -Identity $DLMail -PrimarySmtpAddress "$GroupMailAddress" -ErrorAction SilentlyContinue -ErrorVariable EmailError
            if($EmailError -ne $null)
            {
                Write-Host $EmailError.Exception[1] -ForegroundColor Yellow
            }
            else
            {
                Write-Host "Group mail address updated successfully" -ForegroundColor Green
            }
        }
        $GroupAccessType = Read-Host "Do you want to change your group access type to(Private or Public)"
        if($GroupAccessType -ne "")
        {
            Set-UnifiedGroup -Identity $DLMail -AccessType "$GroupAccessType" -ErrorAction SilentlyContinue -ErrorVariable AccessTypeError
            if($AccessTypeError -ne $null)
            {
                Write-Host $AccessTypeError.Exception[1] -ForegroundColor Yellow
            }
        }
    }
}
function CloseConnection
{
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

#Connect exchange module
Connect_ExchangeOnline
#Check if parameter passed or not
if($DistributionEmailAddress -ne "")
{
    $GetEmailAddress = $DistributionEmailAddress.Split(",")
}
elseif($InputFile -ne "")
{
    try
    {
    $InputFileInfo = Import-Csv -Header "DLMail","AccessType" -path $InputFile 
    $GetEmailAddress = $InputFileInfo.DLMail
    }
    catch
    {
        Write-Host "File not found" -ForegroundColor Red
        CloseConnection
    }
}
else
{
    $GetEmailAddress = Read-Host "Please enter distrbution mail you want to convert to Microsoft 365 group"
    if($GetEmailAddress -ne "")
    {
        $GetEmailAddress = $GetEmailAddress.Split(",")
    }
    else
    {
       Write-Host "You didn't provide any distribution mail" -ForegroundColor Red
       CloseConnection
    }
}
$GetEligibleDLGroup = (Get-EligibleDistributionGroupForMigration).PrimarySmtpAddress
Foreach($DLMail in $GetEmailAddress)
{
    $DLGroup = Get-DistributionGroup| Where-Object{$_.PrimarySmtpAddress -eq $DLMail}
    if($DLGroup -ne $null)
    {
        if($GetEligibleDLGroup -notcontains $DLMail)
        {
            Write-Host "$($DLGroup.DisplayName) is not eligible for convert to Microsoft 365 group" -ForegroundColor Red
            continue
        }
        $GroupMember = (Get-DistributionGroupMember -Identity $DLMail).Name
        Write-Host "`n$($DLGroup.DisplayName) group conversion process started..." -ForegroundColor Magenta
        # Get access type from InputFile
        if($InputFile -ne "")
        {
            $AccessType = $InputFileInfo| Where-Object{$_.DLMail -eq "$DLMail"}
            $AccessType = $AccessType.AccessType
        }
        else
        {
            $AccessType = Read-Host "Access Type (Private or Public)"
        }
        if(($AccessType -ne 'Private') -and ($AccessType -ne 'Public'))
        {
            $AccessType = "Private"
            Write-Host "Access type is wrong. So we can take access type of the group is private."  -ForegroundColor Red
        }
        Create_M365Group
    }
    else
    {
        Write-Host "$DLMail is not found" -ForegroundColor Red
    }
}
CloseConnection
