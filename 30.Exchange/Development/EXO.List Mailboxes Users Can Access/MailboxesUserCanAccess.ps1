<#
=============================================================================================
Name:           List all Exchange Online mailboxes users can access
Version:        2.0
Website:        m365scripts.com
Script by:      M365Scripts Team

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.Exports mailbox rights for all users by default.
2.Allows to generate mailbox rights report for a list of users through input CSV.
3.List mailboxes the users have Send As/Send On Behalf/Full Access permissions.
4.The script can also be executed with MFA enabled account also. 
5.Export report results to CSV file.
6.The script is scheduler-friendly. i.e., credentials are passed as parameters, so worry not!

For detailed script execution: https://m365scripts.com/exchange-online/URL Sluglist-exchange-online-mailboxes-user-has-access-using-powershell



Change Log:

V1.0 (Apr 27, 2022)- Script created
v1.1 (Jun 29, 2022)- Error handling added
v1.3 (July 3, 2024)- Usability added
V2.0 (Jan 09, 2025)- Due to MS update, GUID shown instead of name in the Identity. Updated the script to show name.

============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$UserName = $NULL,
    [string]$Password = $NULL,
    [string]$UPN = $NULL,
    [string]$CSV = $NULL,
    [switch]$FullAccess,
    [switch]$SendAs,
    [switch]$SendOnBehalf
)
function Connect_Exo {
    #Check for EXO v2 module inatallation
    $Module = Get-Module ExchangeOnlineManagement -ListAvailable
    if ($Module.count -eq 0) { 
        Write-Host "Exchange Online PowerShell V2 module is not available"  -ForegroundColor yellow  
        $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No" 
        if ($Confirm -match "[yY]") { 
            Write-host "Installing Exchange Online PowerShell module"
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
        } 
        else { 
            Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
            Exit
        }
    } 
    Write-Host Connecting to Exchange Online...
    #Importing Module by default will avoid the cmdlet unrecognized error 
    Import-Module ExchangeOnline -ErrorAction SilentlyContinue -Force
    #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
    if (($UserName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword
        Connect-ExchangeOnline -Credential $Credential
    }
    else {
        Connect-ExchangeOnline
    }
    Write-Host "ExchangeOnline PowerShell module is connected successfully"
}
function FullAccess {
    $MB_FullAccess = $global:Mailbox | Get-MailboxPermission -User $UPN -ErrorAction SilentlyContinue | Select-Object Identity
    if ($MB_FullAccess.count -ne 0) {
        $ExportResult = @{'User Name' = $Identity; 'AccessType' = "Full Access"; 'Delegated Mailbox Name' = $MB_FullAccess.Identity -join (",") }
    }
    else {
        $ExportResult = @{'User Name' = $Identity; 'AccessType' = "Full Access"; 'Delegated Mailbox Name' = "-" }
    }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'User Name', 'AccessType', 'Delegated Mailbox Name' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
}
function SendAs {
    $MB_SendAs = Get-RecipientPermission -Trustee $UPN -ErrorAction SilentlyContinue | Select-Object Identity
    if ($MB_SendAs.count -ne 0) {
        $ExportResult = @{'User Name' = $Identity; 'AccessType' = "Send As"; 'Delegated Mailbox Name' = $MB_SendAs.Identity -join (",") }
    }
    else {
        $ExportResult = @{'User Name' = $Identity; 'AccessType' = "Send As"; 'Delegated Mailbox Name' = "-" }
    }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'User Name', 'AccessType', 'Delegated Mailbox Name' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
    
}
function SendOnBehalfTo {
    $MB_SendOnBehalfTo = $global:Mailbox | Where-Object { $_.GrantSendOnBehalfTo -match $Identity } -ErrorAction SilentlyContinue | Select-Object Name
    if ($MB_SendOnBehalfTo.count -ne 0) {
        $ExportResult = @{'User Name' = $Identity; 'AccessType' = "Send on Behalf"; 'Delegated Mailbox Name' = $MB_SendOnBehalfTo.Name -join (",") }
    }
    else {
        $ExportResult = @{'User Name' = $Identity; 'AccessType' = "Send on Behalf"; 'Delegated Mailbox Name' = "-" }
    }
    $ExportResults = New-Object PSObject -Property $ExportResult
    $ExportResults | Select-object 'User Name', 'AccessType', 'Delegated Mailbox Name' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force
}


Connect_Exo 
$global:ExportCSVFileName = "MailboxesUserHasAccessTo_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
$global:Mailbox = Get-Mailbox -ResultSize Unlimited
if (($UPN -ne "")) {
    $UserInfo = $global:Mailbox | Where-Object { $_.UserPrincipalName -eq "$UPN" } | Select-Object Identity
    $Identity = $UserInfo.Identity
    if ($FullAccess.IsPresent) {
        FullAccess
    }
    if ($SendAs.IsPresent) {
        SendAs
    }
    if ($SendOnBehalf.IsPresent) {
        SendOnBehalfTo
    }
    if((($FullAccess.IsPresent) -eq $false) -and (($SendAs.IsPresent) -eq $false) -and (($SendOnBehalf.IsPresent) -eq $false)){
        FullAccess
        SendAs
        SendOnBehalfTo
    }
}
elseif (($CSV -ne "")) {
    Import-Csv $CSV -ErrorAction Stop | ForEach-Object {
        $UPN = $_.UPN
        $UserInfo = $global:Mailbox | Where-Object { $_.UserPrincipalName -eq "$UPN" } 
        $Identity = $UserInfo.Displayname
        Write-Progress "Processing for the Mailbox: $Identity"
        if ($FullAccess.IsPresent) {
            FullAccess
        }
        if ($SendAs.IsPresent) {
            SendAs
        }
        if ($SendOnBehalf.IsPresent) {
            SendOnBehalfTo
        }
        if((($FullAccess.IsPresent) -eq $false) -and (($SendAs.IsPresent) -eq $false) -and (($SendOnBehalf.IsPresent) -eq $false)){
            FullAccess
            SendAs
            SendOnBehalfTo
        }
    }
}
else {
    $MBCount = 0
    $global:Mailbox | ForEach-Object {
        $MBCount = $MBCount + 1
        $UPN = $_.UserPrincipalName
        $Identity = $_.DisplayName
        Write-Progress -Activity "Processing for  : $Identity" -Status "Processed mailbox Count: $MBCount" 
        if ($FullAccess.IsPresent) {
            FullAccess
        }
        if ($SendAs.IsPresent) {
            SendAs
        }
        if ($SendOnBehalf.IsPresent) {
            SendOnBehalfTo
        }
        if((($FullAccess.IsPresent) -eq $false) -and (($SendAs.IsPresent) -eq $false) -and (($SendOnBehalf.IsPresent) -eq $false)){
            FullAccess
            SendAs
            SendOnBehalfTo
        }
    }
}
if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") {     
    Write-Host `n "The Output file availble in: " -NoNewline -ForegroundColor Yellow; Write-Host "`"$global:ExportCSVFileName`""`n
    $prompt = New-Object -ComObject wscript.shell    
    $userInput = $prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)    
    if ($userInput -eq 6) {    
        Invoke-Item "$global:ExportCSVFileName"
    }  
}
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
Write-Host "Disconnected active ExchangeOnline session"
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
