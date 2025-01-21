<#
=============================================================================================
Name:           Set up Recurring OOF Replies in Outlook
Description:    This script helps to automate recurring out-of-office replies for Microsoft 365 users
Version:        1.0
Website:        m365scripts.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script can be executed using either username/password or certificate-based authentication (CBA).
2. It automatically calculates and sets the out-of-office start and end times for the upcoming occurrences based on the provided day and time parameters.
3. This script is primarily designed for setting up recurring out-of-office replies with the help of scheduling tools like Windows Task Scheduler or Azure Automation.
4. This script allows admins to set up recurring out-of-office replies for themselves or their users, and it also enables users to configure their own recurring OOF replies.
5. The external out-of-office message defaults to the internal message if not provided.
6. The external audience defaults to "All" if not specified.

For detailed Script execution: https://m365scripts.com/exchange-online/how-to-set-recurring-automatic-replies-in-outlook/
============================================================================================
#>
param (
    [Parameter(Mandatory=$true)][string]$Identity,
    [Parameter(Mandatory=$true)][string]$StartDay,
    [Parameter(Mandatory=$true)][string]$StartTime,
    [Parameter(Mandatory=$true)][string]$EndDay,
    [Parameter(Mandatory=$true)][string]$EndTime,
    [Parameter(Mandatory=$true)][string]$InternalMessage,
    [string]$ExternalMessage,
    [string]$ExternalAudience,
    [string]$UserName,
    [string]$Password,
    [string]$ClientId,
    [string]$Organization,
    [string]$CertificateThumbprint        
)

# Default ExternalMessage and ExternalAudience if not provided
if (-not $ExternalMessage) { $ExternalMessage = $InternalMessage }
if (-not $ExternalAudience) { $ExternalAudience = "All" }

# Function to validate day of the week
function Validate-DayOfWeek {
    param (
        [string]$dayOfWeek
    )
    return [Enum]::GetNames([System.DayOfWeek]) -contains $dayOfWeek
}

# Function to validate time format
function Validate-TimeFormat {
    param (
        [string]$time
    )
    return $time -match '^(0?[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$' # 24-hour format
}

# Validate all parameters at the start
if (-not (Validate-DayOfWeek $StartDay) -or -not (Validate-DayOfWeek $EndDay) -or -not (Validate-TimeFormat $StartTime) -or -not (Validate-TimeFormat $EndTime)) {
    Write-Host "There was an error while calculating the startDateTime or endDateTime. Please check the spelling and format of all the day/time parameters." -ForegroundColor Red
    exit
}

function Get-DateForDayOfWeek ($dayOfWeek) {
    $daysToAdd = ([Enum]::Parse([System.DayOfWeek], $dayOfWeek)) - (Get-Date).DayOfWeek
    if ($daysToAdd -lt 0) { $daysToAdd += 7 }
    (Get-Date).AddDays($daysToAdd).Date
}

# Calculate start and end DateTimes
$startDateTime = (Get-DateForDayOfWeek $StartDay).AddHours([int]$StartTime.Split(':')[0]).AddMinutes([int]$StartTime.Split(':')[1])
$endDateTime = (Get-DateForDayOfWeek $EndDay).AddHours([int]$EndTime.Split(':')[0]).AddMinutes([int]$EndTime.Split(':')[1])
if ($endDateTime -lt $startDateTime) { $endDateTime = $endDateTime.AddDays(7) }

 Write-Host Connecting to Exchange Online...
 #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential
 }
 elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
   Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization
 }

# Set AutoReply configuration
try {
    # Attempt to set the mailbox auto-reply configuration
    Set-MailboxAutoReplyConfiguration –Identity $Identity -AutoReplyState Scheduled –StartTime $startDateTime -EndTime $endDateTime –InternalMessage $InternalMessage -ExternalMessage $ExternalMessage -ExternalAudience $ExternalAudience -ErrorAction Stop
    # If no error occurs, print the success message
    Write-Host "Out-of-office automatic replies configured for the user: $Identity"
}
catch {
    # If an error occurs, print the error message and stop the success message
    Write-Host "Failed to configure out-of-office automatic replies for the user: $Identity" -ForegroundColor Red
    Write-Host $_.Exception.Message
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports ~~" -ForegroundColor Green `n`n
