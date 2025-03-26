<#
====================================================================================================
Name:           Identify and Block Sign-in for Shared Mailboxes in Microsoft 365
Version:        1.0
Website:        o365reports.com

Script Highlights:
~~~~~~~~~~~~~~~~~
1. Export sign-in status report for all the resource and shared mailboxes into a CSV file, including: 
    a. Shared mailboxes – Check sign-in settings for shared accounts. 
    b. Room mailboxes – Track sign-in configurations for meeting rooms. 
    c. Equipment mailboxes – Review sign-in status for assigned equipment. 
2. Blocks sign-in for all shared mailboxes upon confirmation. 
3. Allows to verify sign-in configuration for specific mailbox types and blocks them. 
4. The script automatically verifies and installs the Exchange Online PowerShell and Microsoft Graph Beta modules (if not installed already) upon your confirmation. 
5. Allows users to modify the generated CSV report and provide it as input later to block the sign-in on respective mailboxes. 
6. Provides a detailed log file after disabling sign-in for shared and resource mailboxes. 
7. The script can be executed with an MFA-enabled account too. 
8. The script supports Certificate-based authentication (CBA).

For detailed Script execution: https://o365reports.com/2025/03/25/identify-and-block-sign-in-to-shared-mailbox-using-powershell/
====================================================================================================
#>

param (
    [string] $CertificateThumbPrint,
    [string] $ClientId,
    [string] $Organization,
    [string] $TenantId,
    [Switch] $SharedMailboxOnly,
    [Switch] $RoomMailboxOnly,
    [Switch] $EquipmentMailboxOnly,
    [string] $CSV
)

if($CSV){
    if (-not (Test-Path $CSV -PathType Leaf)) 
    {
        Write-Host "Error: The specified CSV file does not exist or is not accessible." -ForegroundColor Red
        Exit
    }
}

Function ConnectModules
{
    $MgGraphBetaModule =  Get-Module Microsoft.Graph.Beta -ListAvailable
    if($MgGraphBetaModule -eq $null)
    { 
        Write-Host "Important: Microsoft Graph Beta module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $Confirm = Read-Host Are you sure you want to install Microsoft Graph Beta module? [Y] Yes [N] No  
        if($Confirm -match "[yY]") 
        { 
            Write-Host "Installing Microsoft Graph Beta module..."
            try{
                Install-Module Microsoft.Graph.Beta -Scope CurrentUser -AllowClobber
            }
            catch{
                Write-Host "Error occurred : $( $_.Exception.Message )" -ForegroundColor Red
                Exit
            }
            Write-Host "Microsoft Graph Beta module is installed in the machine successfully" -ForegroundColor Magenta 
        } 
        else
        { 
            Write-Host "Exiting. `nNote: Microsoft Graph Beta module must be available in your system to run the script" -ForegroundColor Red
            Exit 
        } 
    }
    $EXOModule=Get-Module ExchangeOnlineManagement -ListAvailable
    if($EXOModule.count -eq 0)
    {
        Write-Host "Important: Exchangeonlinemanagement module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." -ForegroundColor Red
        $Confirm = Read-Host Are you sure want to install module? [Y] Yes [N] No
        if($Confirm -match "[yY]")
        {
            Write-Host Installing Exchange Online Powershell module
            try
            {
                Install-Module -Name ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
            }
            catch
            {
                Write-Host "Error occurred : $( $_.Exception.Message)" -ForegroundColor Red
                Exit
            }
            Write-Host ExchangeOnline installed successfully -ForegroundColor Green
        }
        else
        {
            Write-Host "EXO module is required. Please Install-module ExchangeOnlineManagement."
            Exit
        }
    }

    try
    {
        if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne "") -and ($Organization -ne ""))  
        {   
            Write-Host "Connecting to ExchangeOnline." 
            Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbPrint -AppId $ClientId -Organization $Organization -ErrorAction stop -ShowBanner:$false
            Write-Host "ExchangeOnline connected successfully" -ForegroundColor Green
            Write-Host "Connecting to Microsoft Graph."
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction stop | Out-Null
            Write-Host "Microsoft Graph connected successfully" -ForegroundColor Green
        }
        else
        {
            Write-Host "Connecting to ExchangeOnline"
            Connect-Exchangeonline -ErrorAction stop -ShowBanner:$false
            Write-Host "ExchangeOnline connected successfully" -ForegroundColor Green
            Write-Host "Connecting to Microsoft.Graph."
            #Import-Module Microsoft.Graph.Authentication
            Connect-MgGraph -scopes "Directory.AccessAsUser.All" -ErrorAction Stop | Out-Null
            Write-Host "Microsoft Graph connected successfully" -ForegroundColor Green
        }
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Disconnect-ExchangeOnline -confirm:$false
        Exit
    }
}

Function WriteToLogFile ($message)
{
    $message >> $global:logfile
} 

Function CheckMailboxSigninStatus
{
    param(
        $ExoMailbox 
    )
    $mailbox = Get-MgBetaUser -Userid $ExoMailbox.userprincipalname 
    $DisplayName = $mailbox.DisplayName
    Write-Progress -Activity "`n  Retrieving Signin status for the mailbox: $DisplayName" -status "Mailbox count:$global:RetrieveCount"
    $global:RetrieveCount++
    $UserPrincipalName = $mailbox.UserPrincipalName
    $AccountEnabled = $mailbox.AccountEnabled
    $RecipientTypeDetails = $ExoMailbox.RecipientTypeDetails
    if($RecipientTypeDetails -ne "SharedMailbox" -and $RecipientTypeDetails -ne "RoomMailbox" -and $RecipientTypeDetails -ne "EquipmentMailbox"){
        Write-Host "$userprincipalname is not a shared/room/equipment mailbox" -ForegroundColor Red
        return
    }
    if($AccountEnabled -and !$global:flag){
        $global:flag = $true
    }
    $ExportResult = @{'Display Name' = $DisplayName; 'User Principal Name' = $UserPrincipalName; 'Account Enabled' = $AccountEnabled; 'Recipient Type Details' = $RecipientTypeDetails }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'Display Name', 'User Principal Name', 'Account Enabled', 'Recipient Type Details' | Export-csv -path $global:ExportCSV -NoType -Append -Force
}

Function BlockSignin
{   
    Import-Csv $global:ExportCSV | ForEach-Object{
        if($_.'Account Enabled' -eq $true){
            $Displayname = $_.'Display Name'
            Write-Progress -Activity "Updating signin status for the mailbox: $DisplayName" -status "processing:$global:UpdatedCount"
            $global:UpdatedCount++      
            $UserPrincipalName = $_.'User Principal Name'
            try{
                Update-MgBetaUser -UserId $UserPrincipalName -AccountEnabled:$false -ErrorAction Stop
                WriteToLogFile " Successfully blocked signin for the mailbox: $UserPrincipalName"
            }
            catch{
                WriteToLogFile " Error occured while trying to block the mailbox:  $UserPrincipalName  "
                Write-Host "An error occurred for the mailbox: $UserPrincipalname" ": $( $_.Exception.Message )"
            }
        }
    }
    if(((Test-Path -Path $global:logfile) -eq "True")){
        Write-host "The log file $global:logfile available in $global:Location" -ForegroundColor yellow
    }
}

Function ConfirmationToBlock
{
    Write-Host "`n`nDo you want to proceed with blocking signin enabled mailboxes?  Yes[y] No[n]?"
    $confirm = (Read-Host "Enter your choice").ToUpper()
    if($confirm -eq 'Y')
    {
        Write-Host "Updating Mailbox Signin Status..."
        BlockSignin
    }
}

Function GettingMailbox
{
    param(
        $RecipientTypeDetails
    )
    Get-EXOMailbox -RecipientTypeDetails $RecipientTypeDetails -ResultSize unlimited | ForEach-Object {
        CheckMailboxSigninStatus -ExoMailbox $_
    }
}
Function OutputFileCreation
{
    param(
        $OutputFileName 
    )
    $global:ExportCSV = $OutputFileName + "Mailboxes_SigninStatus_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
    $global:logfile = "DisabledSigninLogFile_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".txt"
}

#........................................................Execution starts here..................................... .........................
#Calling Connection Function
ConnectModules
#To check whether the mailbox have accountenabled true 
$global:flag =$false
#No of mailboxes retrieved count 
$global:RetrieveCount = 1
#No of signin updated count
$global:UpdatedCount = 1

if($SharedMailboxOnly)
{
    $FileName = "Shared"
    OutputFileCreation -OutputFileName $FileName
    GettingMailbox -RecipientTypeDetails SharedMailbox
}
elseif($RoomMailboxOnly)
{
    $FileName = "Room"
    OutputFileCreation -OutputFileName $FileName
    GettingMailbox -RecipientTypeDetails RoomMailbox
}
elseif($EquipmentMailboxOnly)
{
    $FileName = "Equipment"
    OutputFileCreation -OutputFileName $FileName
    GettingMailbox -RecipientTypeDetails EquipmentMailbox
}
elseif($CSV)
{
    $FileName = "Shared_Room_Equipment"
    OutputFileCreation -OutputFileName $FileName
    Import-Csv $CSV | ForEach-object{
        $Mailbox = Get-EXOMailbox $_.'User Principal Name'
        CheckMailboxSigninStatus -ExoMailbox $Mailbox
    }
}
else
{
    $FileName = "SharedandResource"
    OutputFileCreation -OutputFileName $FileName
    GettingMailbox -RecipientTypeDetails SharedMailbox,RoomMailbox,EquipmentMailbox
}

$global:Location = Get-Location

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n

#Invoking outputFile
if (((Test-Path -Path $global:ExportCSV) -ne "True")) {     
    Write-Host "`nNo Mailboxes matches the given criteria." -ForegroundColor Cyan
}
else {
    Write-Host "`nThe output file $global:ExportCSV is available in  the $global:Location" -ForegroundColor Yellow
        $prompt = New-Object -ComObject wscript.shell    
        $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
        if ($userInput -eq 6) {
            Invoke-Item "$global:ExportCSV"
        }
    if($global:flag){
        ConfirmationToBlock
    }
    else{
        Write-Host "Signin status has been already disabled for all the mailboxes" -ForegroundColor Cyan
    }
}

Disconnect-ExchangeOnline -confirm:$false
Disconnect-Mggraph | Out-Null

