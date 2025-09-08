
Param
(
    [Parameter(Mandatory = $false)]
    [string]$UserName = $Null,
    [string]$Password = $Null,
    [int]$Days = -1,
    [string]$UPN = $Null,
    [string]$CSV = $Null
)
Function Connect_Exo {
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
            Write-Host "EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet." 
            Exit
        }
    } 
    Write-Host Connecting to Exchange Online...
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
}
function WriteToLogFile ($message) {
    $message >> $logfile
}
Function DeleteMail {
    $date = (get-date).AddDays( - ($Days)).ToString("MM/dd/yyyy")
    Write-Progress "Deleting mails older than $Days days from $UPN mailbox."
    $DeleteInfo = Search-Mailbox -Identity $UPN -SearchQuery received<=$date  -DeleteContent -Force -ErrorAction Stop -WarningAction SilentlyContinue
    $Identity = $DeleteInfo.Identity
    $ResultItemsCount = $DeleteInfo.ResultItemsCount
    $ResultItemsSize = $DeleteInfo.ResultItemsSize.split("(") | Select-Object -Index 0 
    $Success=$DeleteInfo.Success
    $global:Result = @{'Mailbox Name' = $Identity; 'Deleted mail count' = $ResultItemsCount; 'Deleted mail size' = $ResultItemsSize; 'Result' = $Success }
}

Connect_Exo
Write-Host "Note: Ensure that you have assigned with Mailbox Import Export role." -ForegroundColor Cyan
$global:ExportCSVFileName = "MailDeletionReport_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"
$global:logfile = "MailDeletionLog_" + ((Get-Date -Format "MMM-dd hh-mm-ss tt").ToString()) + ".txt"
if ($UPN -ne "") {
    if (($Days -eq -1)) {
       [int]$Days = Read-Host "Enter number of days"
       try {
            DeleteMail
            $Results = New-Object PSObject -Property $global:Result
            $Results | Format-List 'Mailbox Name', 'Deleted mail count', 'Deleted mail size', 'Success'
        }
        catch {
            Write-Host "Error occured , Please go through your inputs and try again" -ForegroundColor Red
        }
    }
    else {
        try {
            DeleteMail
            $Results = New-Object PSObject -Property $global:Result
            $Results | Format-List 'Mailbox Name', 'Deleted mail count', 'Deleted mail size', 'Success'
        }
        catch {
            Write-Host "Error occured , Please go through your inputs and try again" -ForegroundColor Red
        }
    }
}

elseif ($CSV -ne "") {
    if ($Days -eq -1) {
        [int]$Days = Read-Host "Enter number of days"
        Import-Csv $CSV | ForEach-Object {
            $UPN = $_.UPN
            try {
                DeleteMail
                WriteToLogFile "Deletion process done successfully for $UPN mailbox."
                $Results = New-Object PSObject -Property $global:Result
                $Results | Select-object 'Mailbox Name', 'Deleted Mail Count', 'Deleted Mail Size', 'Result' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force -ErrorAction Stop
            }
            catch {
                WriteToLogFile "Error Occured while deleting mail from $UPN mailbox.Please check the inputs and try again."
            }
        }
    }
    else {
        Import-Csv $CSV | ForEach-Object {
            $UPN = $_.UPN
            try {
                DeleteMail
                WriteToLogFile "Deletion process done successfully for $UPN mailbox."
                $Results = New-Object PSObject -Property $global:Result
                $Results | Select-object 'Mailbox Name',  'Deleted Mail Count', 'Deleted Mail Size', 'Result' | Export-csv -path $global:ExportCSVFileName -NoType -Append -Force -ErrorAction Stop
            }
            catch {
                WriteToLogFile "Error Occured while deleting mail from $UPN mailbox.Please check the inputs and try again."
            }
        }
    }
}

else {
    $UPN = Read-Host "Enter the identity of the user"
    $Days = Read-Host "Enter number of days"
    try {
        DeleteMail
        $Results = New-Object PSObject -Property $global:Result
        $Results | Format-List 'Mailbox Name', 'Deleted Mail Count', 'Deleted Mail Size', 'Result'
    }
    catch {
        Write-Host "Error occured , Please check whether you have necessary permissiosn and go through your inputs" -ForegroundColor Red
    }
}

if ((Test-Path -Path $global:logfile) -eq "True") {
    if ((Test-Path -Path $global:ExportCSVFileName) -eq "True") {
        Write-Host "Deleted email size report availble in `"$global:ExportCSVFileName`"" -ForegroundColor Green 
        Write-Host "The Logfile is availble in $global:logfile."
        $prompt = New-Object -ComObject wscript.shell    
        $userInput = $prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)    
        if ($userInput -eq 6) {    
            Invoke-Item "$global:ExportCSVFileName"
            Invoke-Item "$global:logfile"
        } 
    }
    else {
        Write-Host "The Logfile is availble in $global:logfile."
        $prompt = New-Object -ComObject wscript.shell    
        $userInput = $prompt.popup("Do you want to open output files?", 0, "Open Output File", 4)    
        if ($userInput -eq 6) {
            Invoke-Item "$global:logfile"
        }
    }
}
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
Write-Host "Disconnected active ExchangeOnline session"

<#
=============================================================================================
Name:           Delete older emails in Outlook using PowerShell
Description:    This script deletes emails older than x days using PowerShell and exports log file & report to CSV file
Website:        m365scripts.com
For detailed script execution: https://m365scripts.com/exchange-online/how-to-delete-older-emails-in-outlook-using-powershell
============================================================================================
#>