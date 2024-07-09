 <#
=============================================================================================
Name: Find Inactive Distribution Lists Using PowerShell
Version: 1.0
Website: o365reports.com

~~~~~~~~~~~~~~~~~
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the Exchange PowerShell module (if not installed already) upon your confirmation.   
2. Exports inactive days of distribution lists in Microsoft 365. 
3. Retrieves the last email received date.  
4. Exports report results to CSV.  
5. The script is schedular-friendly.  
6. It can be executed with certificate-based authentication (CBA) too.  

For detailed script execution: https:https://o365reports.com/2024/07/09/find-inactive-distribution-lists-using-powershell/

=============================================================================================
#> 
param(
    [string] $CertificateThumbPrint,
    [string] $ClientId,
    [string] $Organization,
    [string] $UserName,
    [string] $Password,
    [String]$HistoricalMessageTraceReportPath
)

if($HistoricalMessageTraceReportpath)
{
    if (-not (Test-Path $HistoricalMessageTraceReportPath -PathType Leaf)) 
    {
        Write-Host "`nError: The specified CSV file does not exist or is not accessible." -ForegroundColor Red
        Exit
    } 
}
else
{
    Write-Host "`nInput file path is required" -ForegroundColor Red
    Exit
}

Function ConnectEXO
{
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
    Write-Host "`nExchangeonline connected successfully" -ForegroundColor Green
}

Function GettingDistributionLists{
    Get-DistributionGroup -ResultSize unlimited | Where-Object{$_.RecipientTypeDetails -eq "MailUniversalDistributionGroup"} | ForEach-Object{
        ForEach($Email in $_.EmailAddresses){
            $Mail = (($Email).split(":") | Select -Index 1)
            $global:DistributionLists[$Mail] = $_.PrimarySmtpAddress
        }
    }
}

Function FindDifference{
    param(
        $DateTime
    )

    $DayDifference = (New-TimeSpan -Start $DateTime -End (Get-Date)).Days
    $DayDifference = "$DayDifference" + " Days"
    return $DayDifference
}

Function GettingInactiveDistributionLists
{
    # Read the original CSV file content
    $originalContent = Get-Content -Path $HistoricalMessageTraceReportPath
    # Remove non-printable characters (if any)
    $sanitizedContent = $originalContent -replace '[^\P{C}]', ''
    Set-Content -Path $HistoricalMessageTraceReportPath -Value $sanitizedContent

    Import-CSV -path $HistoricalMessageTraceReportPath | Foreach-Object {
        $RecipientAddresses = $_.'recipient_status' -Split ';'
        ForEach($RecipientAddress in $RecipientAddresses){
            $Recipient = $RecipientAddress -Split "##" | select -Index(0)
            $DL = $global:DistributionLists.$Recipient
            if($DL)
            {
                $LastEmail = Get-Date -Date $($_.'origin_timestamp_utc')
                $LastEmailReceived = $LastEmail.ToString("dd-MM-yyyy  HH:mm:ss")                
                $Difference = FindDifference -DateTime $LastEmail
                if($global:InactiveDistributionLists.$DL)
                {
                    $Last = $global:InactiveDistributionLists.$DL[0]
                    if($Last -lt $LastEmailReceived)
                    {
                        $global:InactiveDistributionLists.$DL[0] = $LastEmailReceived
                        $global:InactiveDistributionLists.$DL[1] = $Difference
                    }
                }
                else
                {
                    $global:InactiveDistributionLists[$DL] = @($LastEmailReceived,$Difference)
                }
            }
        }
    }
}

#To connect exchangeonline
ConnectEXO
#Hashmap to store  distribution lists
$global:DistributionLists = @{}
write-Host "`nGetting inactive distribution lists. This may take some time to process based on the input given..."
#To Get all the distribution lists including alias
GettingDistributionLists
#To store the Inactive DistributionLists
$global:InactiveDistributionLists = @{}
GettingInactiveDistributionLists

$global:CSVFile = "InactiveDistributionLists_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"


ForEach($DistributionList in $global:DistributionLists.GetEnumerator()){
    $DL = $($DistributionList.Name) 
    if($DL -eq $($DistributionList.Value)){
        if($global:InactiveDistributionLists.$DL)
        {
            $LastEmailReceivedDate = $global:InactiveDistributionLists.$DL[0]
            $InactiveDays = $global:InactiveDistributionLists.$DL[1]
        }
        else
        {
            $LastEmailReceivedDate = "No data available"
            $InactiveDays = "Inactive for longer than the selected time range."
        }
        $ExportResult = @{'Group Name' =$DL; 'Last Email Received Date' = $LastEmailReceivedDate; 'Inactive Days' = $InactiveDays }
        $ExportResults = New-Object PSObject -Property $ExportResult
        $ExportResults | Select-object 'Group Name', 'Last Email Received Date', 'Inactive Days' | Export-csv -path $global:CSVFile -NoType -Append -Force
    }
}

$Location = Get-Location
if (((Test-Path -Path $global:CSVFile) -eq "True")) {     
    Write-Host "`nThe output file " -NoNewline -ForegroundColor yellow; Write-Host $global:CSVFile -NoNewline -ForegroundColor Cyan; Write-Host " is available in the directory $Location " -ForegroundColor Yellow
    $prompt = New-Object -ComObject wscript.shell    
    $UserInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
    if ($UserInput -eq 6) {
        Invoke-Item "$global:CSVFile"
    }
}

Disconnect-ExchangeOnline -Confirm:$false

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n