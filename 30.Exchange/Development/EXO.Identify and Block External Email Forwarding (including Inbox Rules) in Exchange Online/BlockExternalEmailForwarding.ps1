<#-------------------------------------------------------------------------------------------------------------------------------------------------------------
Name: Identify and Block External Email Forwarding in Exchange Online Using PowerShell
Version: 1.0
Website: o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the Exchange PowerShell module (if not installed already) upon your confirmation. 
2. Exports the ‘External email forwarding report’ and ‘Inbox rules with external forwarding report’ into a CSV file. 
3. Blocks external forwarding configuration for all mailboxes upon confirmation. 
4. Disables all the inbox rules with external forwarding configuration upon confirmation. 
5. Allows to verify external email forwarding for specific mailboxes and blocks them. 
6. Allows users to modify the generated CSV report and provide it as input later to block the respective external forwarding configuration. 
7. Provides the detailed log file after removing the external forwarding configuration and disabling the inbox rules with external forwarding. 
8. The script can be executed with an MFA-enabled account too. 
9. The script supports Certificate-based authentication (CBA).

For detailed script execution: https://o365reports.com/2024/07/23/identify-and-block-external-email-forwarding-in-exo-using-powershell/
--------------------------------------------------------------------------------------------------------------------------------------------------------------#>

param (
    [string] $CertificateThumbPrint,
    [string] $ClientId,
    [string] $Organization,
    [string] $UserName,
    [string] $Password,
    [Switch] $ExcludeGuests,
    [Switch] $ExcludeInternalGuests,
    [String] $MailboxNames,
    [String] $RemoveEmailForwardingFromCSV,
    [String] $DisableInboxRuleFromCSV
)

if($MailboxNames){
    if (-not (Test-Path $MailboxNames -PathType Leaf)) 
    {
        Write-Host "Error: The specified CSV file does not exist or is not accessible." -ForegroundColor Red
        Exit
    }
}
Function WriteToLogFileEmail ($message)
{
    $message >> $global:EmailForwardingLogFile
}
Function WriteToLogFileInboxRule ($message)
{
    $message >> $global:InboxRuleLogFile
}
Function ConnectEXO{
    #check for EXO installation
    $Module=Get-Module ExchangeOnlineManagement -ListAvailable
    if($Module.count -eq 0)
    {
        Write-Host Exchange online powershell is not available -ForegroundColor Yellow
        $Confirm = Read-Host Are you sure want to install module? [Y] Yes [N] No
        if($Confirm -match "[yY]")
        {
            Write-Host Installing Exchange Online Powershell module
            Install-Module -Name ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -scope CurrentUser
            Write-Host ExchangeOnlineManagement installed successfully...
        }
        else
        {
            Write-Host EXO module is required to connect Exchange Online.Please Install-module ExchangeOnlineManagement.
            Exit
        }
    }
    Write-Host "`nConnecting to Exchange Online..."
    try{
        #connect to Exchange Online 
        if(($Organization -ne "") -and ($ClientId -ne "") -and ($CertificateThumbPrint -ne ""))
        {
            #Connect Exchange online using Certificate based Authentication 
            Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbPrint -AppId $ClientId -Organization $Organization -ErrorAction stop -ShowBanner:$false
        }
        elseif(($UserName -ne "") -and ($Password -ne ""))
        {
            #Connect Exchange online using username and password
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword
            Connect-ExchangeOnline -Credential $Credential -ErrorAction stop -ShowBanner:$false
        }
        else
        {
            Connect-ExchangeOnline -ErrorAction stop -ShowBanner:$false
        }
    }
    catch
    {
        Write-Host "Error occurred: $($_.Exception.Message )" -ForegroundColor Red
        Exit
    }
    Write-Host "`nExchangeOnline connected successfully" 
}

Function SplittingLegacyValue{
    param(
        $value
    )
    $SplittedValue = (($value).split(':') | select -Index 1)
    $FinalValue = (($SplittedValue).split(']') | select -Index 0)
    return $FinalValue
}

Function SplittingQuotes{
    param(
        $value
    )
    $SplitValue = (($value).split('"') | Select -Index 1)
    return $SplitValue
}

Function FindForwardingActionIsExternal{
    param(
        $Action
    )
    if($Action.contains("[SMTP:")){
        $SplitQuoteValue = SplittingQuotes -value $Action                   
        return $SplitQuoteValue
    }
    elseif($Action.contains("[EX:")){
        $SplittedLegacy = SplittingLegacyValue -value $Action
        #To Check Internal Guests
        if($ExcludeInternalGuests){
            if($global:InternalGuest.$SplittedLegacy){
                return
            }
        }
        #To check Guest users
        if($global:GuestUsers.$SplittedLegacy){
            if($ExcludeGuests){
                return
            }
            else{
                $SplitQuoteValue = SplittingQuotes -value $Action
                return $SplitQuoteValue
            }
        }
        #To check Mailusers
        if($global:MailUsers.$SplittedLegacy){
            $SplitQuoteValue = SplittingQuotes -value $Action
            return $SplitQuoteValue
        }
        #To check Mailcontacts
        if($global:Contacts.$SplittedLegacy){
            $SplitQuoteValue = SplittingQuotes -value $Action
            return $SplitQuoteValue
        }
    } 
}

Function GetInboxRule{
    param(
        $Mailbox
    )
    Write-Progress -Activity "Getting the inbox rule with external forwarding for the mailbox: $($Mailbox)"
    Get-InboxRule -Mailbox $Mailbox | Where-Object {($_.ForwardAsAttachmentTo -ne $Empty -or $_.ForwardTo -ne $Empty -or $_.RedirectTo -ne $Empty) -and ($_.Enabled -eq $True) } | ForEach-Object{
        $ForwardTo = @()
        $ForwardAsAttachmentTo = @()
        $RedirectTo = @()
        if($_.ForwardTo){
            ForEach($Forward in $_.ForwardTo){
                $IsExternal = FindForwardingActionIsExternal -Action $Forward
                $ForwardTo = $ForwardTo + $IsExternal
            }
        }
        if($_.RedirectTo){
            ForEach($Redirect in $_.RedirectTo){
                $IsExternal = FindForwardingActionIsExternal -Action $Redirect
                $RedirectTo = $RedirectTo + $IsExternal
            }
        }
        if($_.ForwardAsAttachmentTo){
            ForEach($ForwardAsAttach in $_.ForwardAsAttachmentTo){
                $IsExternal = FindForwardingActionIsExternal -Action $ForwardAsAttach
                $ForwardAsAttachmentTo = $ForwardAsAttachmentTo + $IsExternal
            }
        }
        if(($ForwardTo.count -gt 0) -or ($ForwardAsAttachmentTo.count -gt 0) -or ($RedirectTo.count -gt 0)){

            $ExportResult = @{'Mailbox Name' =$_.MailboxOwnerId; 'User Principal Name' = $Mailbox; 'Inbox Rule Name' = $_.Name;'Rule Identity' = $_.Identity;'Forward To' = $ForwardTo -join(","); 'Forward As Attachment To' = $ForwardAsAttachmentTo -join(","); 'Redirect To' = $RedirectTo -join(",")}
            $ExportResults = New-Object PSObject -Property $ExportResult
            $ExportResults | Select-object 'Mailbox Name', 'User Principal Name', 'Inbox Rule Name', 'Rule Identity', 'Forward To', 'Forward As Attachment To', 'Redirect To' | Export-csv -path $global:ExportInboxRule -NoType -Append -Force
        }
    }
    Write-Progress -Activity "Getting the inbox rule with external forwarding for the mailbox: $($Mailbox)" -Completed
}
Function CheckEmailForwardingForExternal{
    param(
        $Mailbox
    )
    Write-Progress -Activity "Getting External Forwarding for the Mailbox: $($Mailbox.DisplayName)"
    $ForwardingAddress = '-'
    $ForwardingSMTPAddress = '-'
    $ExternalForwardingAddress = '-'
    $ExternalForwardingSMTPAddress = '-'
    $ExternalAddressFound = $True
    $ExternalSMTPAddressFound = $True
    #Check whether forwardingaddress is in mailcontact
    if($Mailbox.ForwardingAddress){
        $ExternalAddressFound = $False
        $ForwardingAddress = $Mailbox.ForwardingAddress
        if($global:Contacts.$ForwardingAddress){
            $ExternalForwardingAddress = $ForwardingAddress
            $ExternalAddressFound = $True
        }
    }
    #check whether forwarding smtp address is external domain 
    if($Mailbox.ForwardingSMTPAddress){
        $ExternalSMTPAddressFound = $False
        $ForwardingSMTPAddress = (($Mailbox.ForwardingSMTPAddress).split(":") | Select -Index 1)
        $checkDomainIsInternal = (($Mailbox.ForwardingSMTPAddress).split("@") | Select -Index 1)
        if(!$global:Domain.$checkDomainIsInternal)
        {
            $ExternalForwardingSMTPAddress = $ForwardingSMTPAddress
            $ExternalSMTPAddressFound = $True 
        } 
    }
    #To check Internal Guest Users
    if($ExcludeInternalGuests){
        if(!$ExternalAddressFound){
            if($global:InternalGuest.$ForwardingAddress){
                $ExternalAddressFound = $True
            }
        }
        if(!$ExternalSMTPAddressFound){
            if($global:InternalGuest.$ForwardingSMTPAddress){
                $ExternalSMTPAddressFound = $True
            }
        }
    }
    #To check Guest Users
    if(!$ExternalAddressFound -or !$ExternalSMTPAddressFound){
        if(!$ExternalAddressFound){
            if($global:GuestUsers.$ForwardingAddress){
                if($ExcludeGuests){
                    $ExternalAddressFound = $True
                }
                else{
                    $ExternalForwardingAddress = $ForwardingAddress
                    $ExternalAddressFound = $True
                }
            }
        }
        if(!$ExternalSMTPAddressFound){
            if($global:GuestUsers.$ForwardingSMTPAddress){
                if($ExcludeGuests){
                    $ExternalSMTPAddressFound = $True
                }
                else{
                    $ExternalForwardingSMTPAddress = $ForwardingSMTPAddress
                    $ExternalSMTPAddressFound = $True
                }
            }
        }
    }

    #To Check MailUsers
    if(!$ExternalAddressFound -or !$ExternalSMTPAddressFound){
        if(!$ExternalAddressFound){
            if($global:MailUsers.$ForwardingAddress){
                $ExternalForwardingAddress = $ForwardingAddress
                $ExternalAddressFound = $True
            }
        }
        if(!$ExternalSMTPAddressFound){
            if($global:MailUsers.$ForwardingSMTPAddress){
                $ExternalForwardingSMTPAddress = $ForwardingSMTPAddress
                $ExternalSMTPAddressFound = $True
            }
        }
    }
        
    if(($ExternalForwardingAddress -ne "-") -or ($ExternalForwardingSMTPAddress -ne "-")){
        $ExportResult = @{'Display Name' =$Mailbox.DisplayName; 'User Principal Name' = $Mailbox.UserPrincipalName; 'Forwarding Address' = $ExternalForwardingAddress; 'Forwarding SMTP Address' = $ExternalForwardingSMTPAddress; 'Deliver To Mailbox and Forward' = $Mailbox.DeliverToMailboxAndForward }
        $ExportResults = New-Object PSObject -Property $ExportResult
        $ExportResults | Select-object 'Display Name', 'User Principal Name', 'Forwarding Address', 'Forwarding SMTP Address', 'Deliver To Mailbox and Forward' | Export-csv -path $global:ExportEmailForwarding -NoType -Append -Force
    }
    Write-Progress -Activity "Getting External Forwarding for the Mailbox: $($Mailbox.DisplayName)" -Completed
}

Function GetMailbox{
    if($MailboxNames){
        Import-Csv $MailboxNames | ForEach-object {
            $Mailbox = Get-EXOMailbox $($_.'User Principal Name') -properties ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxAndForward | Where-Object {($_.ForwardingAddress -ne $null) -or ($_.ForwardingSMTPAddress -ne $null)} 
            if($Mailbox) {
                CheckEmailForwardingForExternal -Mailbox $Mailbox
            }
            GetInboxRule -Mailbox $_.'User Principal Name'
        }
    }
    else{
        Get-EXOMailbox -ResultSize unlimited -properties ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxAndForward | ForEach-Object {
            if(($_.ForwardingAddress -ne $null) -or ($_.ForwardingSMTPAddress -ne $null)){
                CheckEmailForwardingForExternal -Mailbox $_
            }
            GetInboxRule -Mailbox $_.UserPrincipalName
        }
    }
}

Function RemoveEmailForwarding{
    param(
        $CSV
    )
    Write-Host "`nRemoving external forwarding..."
    Import-Csv $CSV | ForEach-Object {
        Write-Progress -Activity "Removing external forwarding for the mailbox: $($_.'Display Name')"
        try {
            if($_.'Forwarding Address' -ne '-'){
                if($_.'Forwarding SMTP Address' -ne '-'){
                    Set-Mailbox $_.'User Principal Name' -ForwardingAddress $NULL -ForwardingSMTPAddress $NULL -ErrorAction Stop
                    WriteToLogFileEmail "External forwarding successfully removed from $($_.'Display Name') : ForwardingAddress - $($_.'Forwarding Address') , ForwardingSMTPAddress: $($_.'Forwarding SMTP Address')" 
                }
                else{
                    Set-Mailbox $_.'User Principal Name' -ForwardingAddress $Null -ErrorAction Stop
                    WriteToLogFileEmail "External forwarding successfully removed from $($_.'Display Name') : ForwardingAddress - $($_.'Forwarding Address')"
                }
            }
            else{
                Set-Mailbox $_.'User Principal Name' -ForwardingSMTPAddress $NULL -ErrorAction Stop
                WriteToLogFileEmail "External forwarding successfully removed from $($_.'Display Name') : ForwardingSMTPAddress - $($_.'Forwarding SMTP Address')" 
            }
        }                 
        catch {
            Write-Host "Error occured while removing external forwarding configuration for $($_.'Display Name') : $($_.Exception.Message) ." -ForegroundColor Red
            WriteToLogFileEmail "Error Occured: while removing external forwarding configuration for : $($_.'Display Name')"
        }
    }
}
Function DisableInboxRule{ 
    param(
        $CSV
    )
    Write-Host "`nDisabling inbox rule..."
    Import-Csv $CSV | ForEach-Object{
        Write-Progress -Activity "Disabling inbox rule with external forwarding for the mailbox: $($_.'User Principal Name') "
        try
        {
            Disable-InboxRule -Identity $_.'Rule Identity' -ErrorAction Stop -Confirm:$false
            WriteToLogFileInboxRule "The Inbox rule with external forwarding :  '$($_.'Inbox Rule Name')'  present in : $($_.'Mailbox Name') mailbox is disabled."
        }
        catch
        { 
            Write-Host "Error occured, while processing the inbox rule : $( $_.Exception.Message)" -ForegroundColor Red
            WriteToLogFileInboxRule "Error Occured: While disabling inbox rule with external forwarding : $($_.'Inbox Rule Name') present in $($_.'Mailbox Name')"
        }
    }
}

Function ConfirmationToBlock{
    $confirm = Read-Host "Enter your choice"
    if($confirm -match "[yY]")
    {
        if((Test-Path -path $global:ExportEmailForwarding) -eq "True"){
            RemoveEmailForwarding -CSV $global:ExportEmailForwarding
        }
        if((Test-Path -Path $global:ExportInboxRule) -eq "True"){
            DisableInboxRule -CSV $global:ExportInboxRule
        }
    }
}

Function InvokeOutputFiles{
    if (((Test-Path -Path $global:ExportEmailForwarding) -ne "True") -and ((Test-Path -Path $global:ExportInboxRule) -ne "True") ) {     
        Write-Host "`nExternal forwarding is not enabled for the given input." -ForegroundColor Green  
        Exit
    }
    else {
        if (((Test-Path -Path $global:ExportEmailForwarding) -eq "True") -and ((Test-Path -Path $global:ExportInboxRule) -eq "True")) {
                Write-Host "`nFollowing output files are generated and available in the directory $($global:Location):" -NoNewline -ForegroundColor Yellow
                Write-Host "$global:ExportEmailForwarding , $global:ExportInboxRule" -ForegroundColor Cyan
                $prompt = New-Object -ComObject wscript.shell    
                $userInput = $prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)    
                if ($userInput -eq 6) {    
                    Invoke-Item "$global:ExportEmailForwarding"
                    Invoke-Item "$global:ExportInboxRule"                    
                }
                Write-Host "`nDo you want to proceed with remove external forwarding and disable inbox rule with external forwarding? Yes[Y] No[N]"
                ConfirmationToBlock
            }
            elseif((Test-Path -Path $global:ExportEmailForwarding) -eq "True") {
                Write-Host "`nThe Output file available in the directory $($global:Location):"  -NoNewline -ForegroundColor Yellow
                Write-Host $global:ExportEmailForwarding -ForegroundColor Cyan
                $prompt = New-Object -ComObject wscript.shell    
                $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
                if ($userInput -eq 6) {    
                    Invoke-Item "$global:ExportEmailForwarding"                    
                }
                Write-Host "`nDo you want to proceed with remove external forwarding? Yes[Y] No[N]"
                ConfirmationToBlock
            }
            else {
                Write-Host "`nThe output file available in the directory $($global:Location):" -NoNewline -ForegroundColor Yellow
                Write-Host "$global:ExportInboxRule" -ForegroundColor Cyan
                $prompt = New-Object -ComObject wscript.shell
                $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)
                if ($userInput -eq 6) {
                    Invoke-Item "$global:ExportInboxRule"
                }
                Write-Host "`nDo you want to proceed with disable inbox rule with external forwarding? Yes[Y] No[N]"
                ConfirmationToBlock
            }
    }
}

#................................................Execution starts here.....................................................
#Calling Connection Function
ConnectEXO

if(!$RemoveEmailForwardingFromCSV -and !$DisableInboxRuleFromCSV){
    Write-Host "`nFetching mailboxes whose emails are forwarded to external."
    #Fetching domains
    $global:Domain = @{}
    Get-AcceptedDomain | ForEach-Object{
        $global:Domain[$_.DomainName] = $_.DomainName
    }
    #Fetching Guests 
    $global:GuestUsers = @{}
    Get-User -resultsize unlimited | where-object {$_.UserType -eq 'Guest'} | select Identity,LegacyExchangeDN,UserPrincipalName | ForEach-Object{
        $global:GuestUsers[$_.Identity] = $_.Identity
        $global:GuestUsers[$_.LegacyExchangeDN] = $_.LegacyExchangeDN
        $global:GuestUsers[$_.UserPrincipalName] = $_.UserPrincipalName
    }
    #FetchingInternalGuests
    if($ExcludeInternalGuests){
        $global:InternalGuest = @{}
        Get-User -resultsize unlimited | Where-Object {$_.UserPersona -eq 'InternalGuest'} | Select Identity,LegacyExchangeDN,UserPrincipalName | ForEach-Object{
            $global:InternalGuest[$_.Identity] = $_.Identity
            $global:InternalGuest[$_.LegacyExchangeDN] = $_.LegacyExchangeDN
            $global:InternalGuest[$_.UserPrincipalName] = $_.UserPrincipalName
        }
    }
    #Fetching contacts
    $global:Contacts = @{}
    Get-MailContact -resultsize unlimited | select Identity,LegacyExchangeDN | ForEach-Object{
        $global:Contacts[$_.Identity] = $_.Identity
        $global:Contacts[$_.LegacyExchangeDN] = $_.LegacyExchangeDN
    }
    #Mailusers
    $global:MailUsers = @{}
    Get-MailUser -resultsize unlimited | select Identity,LegacyExchangeDN,UserPrincipalName | ForEach-Object{
        $global:MailUsers[$_.Identity] = $_.Identity
        $global:MailUsers[$_.LegacyExchangeDN] = $_.LegacyExchangeDN
        $global:MailUsers[$_.UserPrincipalName] = $_.UserPrincipalName
    }
    #OutputFiles 
    $global:ExportEmailForwarding = "ExternalEmailForwardingReport_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
    $global:ExportInboxRule = "InboxRulesWithExternalForwardingReport_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
}
$global:EmailForwardingLogFile = "RemovedExternalForwarding_LogFile_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".txt"
$global:InboxRuleLogFile = "DisableInboxRuleWithExternalForwarding_LogFile_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".txt"
$global:Location = Get-Location

if(!$RemoveEmailForwardingFromCSV -and !$DisableInboxRuleFromCSV){
    GetMailbox
    InvokeOutputFiles
}

if($RemoveEmailForwardingFromCSV){
    if((Test-Path -path $RemoveEmailForwardingFromCSV) -eq "True"){
        RemoveEmailForwarding -CSV $RemoveEmailForwardingFromCSV
    }
    else{
        Write-Host "`nError: The specified CSV file for removing external forwarding" -NoNewline -ForegroundColor Red
        Write-Host " $RemoveEmailForwardingFromCSV " -Nonewline 
        Write-Host "is not valid. Please check the path." -ForegroundColor Red
    }
}
if($DisableInboxRuleFromCSV){
    if((Test-Path -path $DisableInboxRuleFromCSV) -eq "True"){
        DisableInboxRule -CSV $DisableInboxRuleFromCSV
    }
    else{
        Write-Host "`nError: The specified CSV file for disabling inbox rule" -NoNewline -ForegroundColor Red
        Write-Host " $DisableInboxRuleFromCSV " -Nonewline 
        Write-Host "is not valid. Please check the path." -ForegroundColor Red
    }
}

if (((Test-Path -Path $global:EmailForwardingLogFile) -eq "True")){
    Write-host "`nThe log file $global:EmailForwardingLogFile available in the directory $global:Location" -ForegroundColor yellow
}


if (((Test-Path -Path $global:InboxRuleLogFile) -eq "True")){
    Write-host "`nThe log file $global:InboxRuleLogFile available in the directory $global:Location" -ForegroundColor yellow
}

Disconnect-ExchangeOnline -confirm:$false
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n