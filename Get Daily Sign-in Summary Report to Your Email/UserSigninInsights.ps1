<#
=============================================================================================
Name:           Send Automated Microsoft 365 User Sign-in Summary Email using PowerShell 
Description:    Automatically sends daily Microsoft 365 user sign-ins summary email that includes a CSV report and an easy-to-read HTML dashboard for quick assessment.
Version:        1.0
Website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Automatically sends daily Microsoft 365 user sign-in summary email to specified recipients. 
2. Installs the Microsoft Graph PowerShell module automatically, if it is not already installed. 
3. Exports detailed user sign-in logs to a CSV file and stores it on your local machine. 
4. Generates a one-view HTML dashboard showing a clear summary of user sign-ins. 
5. Supports certificate-based authentication (CBA) too. 
6. The script is scheduler-friendly. 

For detailed script execution:https://o365reports.com/2025/09/30/automate-microsoft-365-user-sign-in-summary-email-using-powershell/
 
============================================================================================
#>
Param
(
    [switch]$CreateSession,
    [string]$Recipients,
    [Switch]$HideSummaryAtEnd,
    [string]$FromAddress,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

Function Connect_MgGraph
{
 #Check for module installation
 $Module=Get-Module -Name microsoft.graph.beta -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Microsoft Graph PowerShell SDK is not available  -ForegroundColor yellow  
  $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Microsoft Graph PowerShell module..."
   Install-Module Microsoft.Graph.beta -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
  }
  else
  {
   Write-Host "Microsoft Graph Beta PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph cmdlet." 
   Exit
  }
 }
 #Disconnect Existing MgGraph session
 if($CreateSession.IsPresent)
 {
  Disconnect-MgGraph
 }


 Write-Host Connecting to Microsoft Graph...
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
 }
 else
 {
  Connect-MgGraph -Scopes "AuditLog.Read.All", "Directory.Read.All", "Policy.Read.ConditionalAccess", "Mail.Send"  -NoWelcome
 }
}


Function SendEmail 
{
 #Recipients Address handling
 $EmailAddresses = ($Recipients -split ",").Trim()
 $ToRecipients = @()
 foreach ($Email in $EmailAddresses) {
  $ToRecipients += @{
   emailAddress = @{
    address = $Email
   }
  }
 }

 #Read the HTML as attachment content
 $FileBytes = [System.IO.File]::ReadAllBytes($ExportHTML)         
 $HtmlBase64 = [System.Convert]::ToBase64String($FileBytes)    

 # Read the CSV file
 $CsvFileBytes = [System.IO.File]::ReadAllBytes($ExportCSV)         
 $CsvBase64 = [System.Convert]::ToBase64String($CsvFileBytes)    

  $EmailBody = @"
 <html>
  <head>
    <meta charset='UTF-8'>
  </head>
  <body>
    <p>Hello Admin,</p>
    <p>Here is the daily summary report for the date $EndDate.</p>


    <p>If you notice any unusual activity or need to investigate a specific entry in the HTML report, please check CSV report. You can also follow up via 
    <a href='https://entra.microsoft.com/#view/Microsoft_AAD_IAM/SignInLogsList.ReactView/timeRangeType/last24hours/showApplicationSignIns~/true' target='_blank'>Entra signin logs</a>.</p>  

    <p>If you have any questions, feel free to contact IT support.</p>  
    <p>Best regards,<br>IT Admin Team</p>
  </body>
 </html>
"@

 $params = @{
        Message = @{
            Subject = "Microsoft 365 Users' Daily Sign-in Insights"
            Body = @{
                ContentType = "HTML"
                Content     = $EmailBody
            }
            
            Attachments = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                Name          = "User SignIn Summary Insights.html"
                ContentBytes  = $HtmlBase64
                ContentType   = "text/html"
            },
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                Name          = "User SignIn Report.csv"
                ContentBytes  = $CsvBase64
                ContentType   = "text/csv"
             }
        )  
            ToRecipients = $ToRecipients
        }
        SaveToSentItems = $true
    }
 Send-MgBetaUserMail -UserId $Script:FromAddress -BodyParameter $params
}
Connect_MgGraph

#Initialization
$Location=Get-Location
$ExportCSV = "$Location\M365Users_Signin_Report$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportHTML= "$Location\M365Users_Signin_SummaryReport$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).html"
$ExportResult=""   
$ExportResults=@()  
$Count=0
$SuccessfulSigninCount=0
$FailedSigninCount=0
$MFASigninCount=0
$NonMFASigninCount=0
$SigninsBlockedByCACount=0
$SigninsGrantedByCACount=0
$ExternalUserSigninCount=0
$ExternalUserSuccessfulSigninCount=0
$ExternalUserFailedSigninCount=0
$FailedSigninUsers=@{}
$CABlockedUsers=@{}
$CAGrantedUsers=@{}
$SuccessfulNonMFSignInUsers=@{}
$ExternalUserSignIns=@{}

$EndDate=(Get-Date).Date
$StartDate=$EndDate.AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ssZ')
$EndDate=$EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
$Filter="createdDatetime ge $StartDate and createdDatetime le $Enddate"

#retrieve signin activities
Write-Host "Generating M365 users' signin report..."
Get-MgBetaAuditLogSignIn -All -Filter $Filter | ForEach-Object {
 $Count++
 $UPN=$_.UserPrincipalName
 Write-Progress -Activity "`n     Processed sign-in record: $count "
 $CreatedDate= (Get-Date($_.CreatedDateTime)).ToLocalTime()  #$_.CreatedDateTime Uncomment to view the Activity Time in UTC
 $Id=$_.Id
 $UserDisplayName=$_.UserDisplayName
 $UPN=$_.UserPrincipalName
 $AuthenticationRequirement=$_.AuthenticationRequirement
 $Location="$($_.Location.City),$($_.Location.State),$($_.Location.CountryOrRegion)"
 $DeviceName=$_.DeviceDetail.DisplayName
 $Browser=$_.DeviceDetail.Browser
 $OperatingSystem=$_.DeviceDetail.OperatingSystem
 $IpAddress=$_.IpAddress
 $ErrorCode=$_.Status.ErrorCode
 $FailureReason=$_.Status.FailureReason
 $UserType=$_.UserType
 $RiskDetail=$_.RiskDetail
 $IsInteractive=$_.IsInteractive
 $RiskState=$_.RiskState
 $AppDispalyName=$_.AppDisplayName
 $ResourceDisplayName=$_.ResourceDisplayName
 $ConditionalAccessStatus=$_.ConditionalAccessStatus
 $AppliedConditionalAccessPolicies=$_.AppliedConditionalAccessPolicies
 


 $AppliedPolicies = @()
 $AppliedConditionalAccessPolicies | ForEach-Object {
    if($_.Result -eq 'Success' -or $_.Result -eq 'Failed'){
        $AppliedPolicies += $_.DisplayName
    }
 }
 if($AppliedPolicies.Count -eq 0){
 $AppliedPolicies = "None"
 } else{
 $AppliedPolicies = $AppliedPolicies -join ", "
 }

 #Signin Status count handling
 if($ErrorCode -eq 0)
 {
  $Status='Success'
  $SuccessfulSigninCount++

  #Authentication requirement filtering for successful sign-in attempts
  if($AuthenticationRequirement -eq "singleFactorAuthentication")
  {
   $NonMFASigninCount++
   if($SuccessfulNonMFSignInUsers.ContainsKey($UPN))
   {
    $SuccessfulNonMFSignInUsers[$UPN].Count++
   }
   else
   {
    $SuccessfulNonMFSignInUsers[$UPN]= @{
    Count=1
    LastAccessedTime=$CreatedDate }
   }
  }
  elseif($AuthenticationRequirement -eq "multiFactorAuthentication")
  {
   $MFASigninCount++
  }
 }
 else
  {
   $Status='Failed'
   $FailedSigninCount++
   if ($FailedSigninUsers.ContainsKey($UPN)) 
   {
    $FailedSigninUsers[$UPN]++     
   } 
   else 
   {
    $FailedSigninUsers[$UPN] = 1
   }
  }
 


 #Finding User signins blocked by CA policies
 if($ConditionalAccessStatus -eq "Failure" -or ($AuthenticationRequirement -eq "singleFactorAuthentication"))
 {
  $SigninsBlockedByCACount++
  if($CABlockedUsers.ContainsKey($UPN))
  {
   $CABlockedUsers[$UPN].Count++
   #$CABlockedUsers[$UPN].LastAccessedTime=$CreatedDate
  }
  else
  {
   $CABlockedUsers[$UPN]= @{
   Count=1
   LastAccessedTime=$CreatedDate }
  }
 }


 #Finding User signins granted by CA policies
 if($ConditionalAccessStatus -eq "Success" -and ($AuthenticationRequirement -eq "multiFactorAuthentication"))
 {
  $SigninsGrantedByCACount++
  if($CAGrantedUsers.ContainsKey($UPN))
  {
   $CAGrantedUsers[$UPN].Count++
   #$CABlockedUsers[$UPN].LastAccessedTime=$CreatedDate
  }
  else
  {
   $CAGrantedUsers[$UPN]= @{
   Count=1
   LastAccessedTime=$CreatedDate }
  }
 }

 #Finding external user sign-in attempts
 if($UserType -ne 'member')
 {
  $ExternalUserSigninCount++
  if(!$ExternalUserSignIns.Contains($UPN))
  {
    $ExternalUserSignIns[$UPN]=@{
     ExtSuccessfulSigninCount=0
     ExtFailedSigninCount=0
     LastAccessedTime=$CreatedDate
     LastSignInStatus="NotDefined"}
  }
  if($ErrorCode -eq 0)
  {
   $ExternalUserSuccessfulSigninCount++
   $ExternalUserSignIns[$UPN].ExtSuccessfulSigninCount++
   if($ExternalUserSignIns[$UPN].LastSigninStatus -eq "NotDefined")
   {
    $ExternalUserSignIns[$UPN].LastSigninStatus = "Success"
   }
  }
  else
  {
   $ExternalUserFailedSigninCount++
   $ExternalUserSignIns[$UPN].ExtFailedSigninCount++
   if($ExternalUserSignIns[$UPN].LastSigninStatus -eq "NotDefined")
   {
    $ExternalUserSignIns[$UPN].LastSigninStatus = "Failed"
   }
  }
 }

 
 if($FailureReason -eq 'Other.') {
   $FailureReason="None"
 }


  #Export users to output file

  $ExportResult=[PSCustomObject]@{'Signin Date'=$CreatedDate; 'User Name'=$UserDisplayName; 'SigninId'=$Id;'UPN'=$UPN; 'Status'=$Status; 'Ip Address'=$IpAddress; 'Location'=$Location; 'Device Name'=$DeviceName; 'Browser'=$Browser; 'Operating System'=$OperatingSystem; 'User type'=$UserType; 'Authentication Requirement'=$AuthenticationRequirement; 'Risk detail'=$RiskDetail; 'Risk state'=$RiskState; 'Conditional access status'=$ConditionalAccessStatus;  'Applied Conditional Access Policies'=$AppliedPolicies; 'IsInteractive'=$IsInteractive;}
  $ExportResult | Export-Csv -Path $ExportCSV -Notype -Append
 
}

#Disconnect the session after execution
#Disconnect-MgGraph | Out-Null


$SortedFailedUsers = $FailedSigninUsers.GetEnumerator() | Sort-Object Value -Descending
$SortedCABlockedUsers = $CABlockedUsers.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending
$SortedCAGrantedUsers = $CAGrantedUsers.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending
$SortedSuccessfulNonMFSignInUsers = $SuccessfulNonMFSignInUsers.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending
$SortedExternalUserSignins=$ExternalUserSignIns.GetEnumerator() | Sort-Object { $_.Value.ExtFailedSigninCount } -Descending

$HtmlContent = @"
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f9f9f9;
            padding: 20px;
        }
        .card-container {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }
        .card {
            background-color: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0px 0px 8px rgba(0, 0, 0, 0.1);
            flex: 1;
            min-width: 200px;
        }
        .card h2 {
            margin: 0;
            font-size: 18px;
            color: #555;
        }
        .card p {
            margin: 10px 0 0 0;
            font-size: 24px;
            font-weight: bold;
            color: #333;
        }
       
        table {
            width: 50%;
            border-collapse: collapse;
            margin-top: 20px;
            margin-bottom: 30px;
        }
        th, td {
            border: 1px solid #ccc;
            padding: 10px;
            text-align: left;
        }
        th {
            background-color: #0078D7;
            color: white;
        }
        caption {
            caption-side: top;
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 10px;
            text-align: left;
        }
        .sub-card-container {
            display: flex;
            gap: 10px;
            margin-top: 15px;
            flex-wrap: wrap;
            max-width: 360px;
        }
        .sub-card {
            background-color: #f1f1f1;
            padding: 10px;
            border-radius: 8px;
            flex: 1 ;
    min-width: 140px;
            box-shadow: inset 0 0 3px rgba(0,0,0,0.05);
        }
        .sub-card h4 {
            margin: 0;
            font-size: 14px;
            color: #666;
        }
        .sub-card p {
            margin: 5px 0 0 0;
            font-size: 20px;
            font-weight: bold;
            color: #222;
        }
    </style>
</head>
<body>
    <h1>Daily User Sign-in Summary Report</h1>
    <div class="card-container">
        <div class="card">
            <h2>Total Sign-ins</h2>
            <p>$Count</p>
        </div>
        <div class="card">
            <h2>Successful Sign-ins</h2>
            <p>$SuccessfulSigninCount</p>
        </div>
        <div class="card">
            <h2>Failed Sign-ins</h2>
            <p>$FailedSigninCount</p>
        </div>
        <div class="card">
            <h2>Blocked by CA</h2>
            <p>$SigninsBlockedByCACount</p>
        </div>
        <div class="card">
            <h2>Granted by CA</h2>
            <p>$SigninsGrantedByCACount</p>
        </div>
        <div class="card">
            <h2>Successful Sign-ins</h2>
         
            <div class="sub-card-container">
                <div class="sub-card">
                    <h4>MFA Sign-ins</h4>
                    <p>$MFASigninCount</p>
                </div>
                <div class="sub-card">
                    <h4>Non-MFA Sign-ins</h4>
                    <p>$NonMFASigninCount</p>
                </div>
            </div>
        </div>
        <div class="card">
            <h2>External User Sign-ins</h2>
         
            <div class="sub-card-container">
                <div class="sub-card">
                    <h4>Successful</h4>
                    <p>$ExternalUserSuccessfulSigninCount</p>
                </div>
                <div class="sub-card">
                    <h4>Failed</h4>
                    <p>$ExternalUserFailedSigninCount</p>
                </div>
            </div>
        </div>

    </div>
"@
# Add Failed Sign-in user table if there are any entries
 if ($FailedSigninUsers.Count -gt 0) {
    $HtmlContent += @"
    <table>
        <caption>Failed Sign-in Users</caption>
        <tr><th>User Principal Name</th><th>Failed Sign-in Count</th></tr>
"@

    foreach ($user in $SortedFailedUsers) {
        $HtmlContent += "<tr><td>$($user.Key)</td><td>$($user.Value)</td></tr>`n"
    }

    $HtmlContent += "</table>`n`n"
    
}

#Add sign-ins blocked by CA policies if there are any entries
 if ($SigninsBlockedByCACount.Count -gt 0) {
    $HtmlContent += @"
    <table>
        <caption>Sign-ins Blocked by CA Policies</caption>
        <tr><th>User Principal Name</th><th>Sign-in Count</th><th>Last Blocked Time</th></tr>
"@

    foreach ($user in $SortedCABlockedUsers) {
        $HtmlContent += "<tr><td>$($user.Key)</td><td>$($user.Value.Count)</td><td>$($user.Value.LastAccessedTime)</td></tr>`n"
    }

    $HtmlContent += "</table>`n"
}

#External user sign-in details if there are entries
 if ($ExternalUserSigninCount.Count -gt 0) {
    $HtmlContent += @"
    <table>
        <caption>External User Sign-in Details</caption>
        <tr><th>User Principal Name</th><th>Successful Sign-in Count</th><th>Failed Sign-in Count</th><th>Last Sign-in Time</th><th>Last Sign-in Status</th></tr>
"@

    foreach ($user in $SortedExternalUserSignins) {
        $HtmlContent += "<tr><td>$($user.Key)</td><td>$($user.Value.ExtSuccessfulSigninCount)</td><td>$($user.Value.ExtFailedSigninCount)</td><td>$($user.Value.LastAccessedTime)</td><td>$($user.Value.LastSignInStatus)</td></tr>`n"
    }

    $HtmlContent += "</table>`n"
}


#Add successful sign-in users using single factor authentication (Non-MFA signins) if there are entries
if ($NonMFASigninCount.Count -gt 0) {
    $HtmlContent += @"
    <table>
        <caption>Successful Sign-ins Using Single Factor Authentication</caption>
        <tr><th>User Principal Name</th><th>Sign-in Count</th><th>Last Sign-in Time</th></tr>
"@

    foreach ($user in $SortedSuccessfulNonMFSignInUsers) {
        $HtmlContent += "<tr><td>$($user.Key)</td><td>$($user.Value.Count)</td><td>$($user.Value.LastAccessedTime)</td></tr>`n"
    }

    $HtmlContent += "</table>`n"
}

#Add sign-ins granted by CA policies if there are any entries
 if ($SigninsGrantedByCACount.Count -gt 0) {
    $HtmlContent += @"
    <table>
        <caption>Sign-ins Granted by CA Policies</caption>
        <tr><th>User Principal Name</th><th>Sign-in Count</th><th>Last Sign-in Time</th></tr>
"@

    foreach ($user in $SortedCAGrantedUsers) {
        $HtmlContent += "<tr><td>$($user.Key)</td><td>$($user.Value.Count)</td><td>$($user.Value.LastAccessedTime)</td></tr>`n"
    }

    $HtmlContent += "</table>`n"
}

# Close HTML tags
$HtmlContent += @"
</body>
</html>
"@

# Write to HTML file
$HtmlContent | Out-File -FilePath $ExportHTML -Encoding UTF8

#Send email if receipient address present
if($Recipients -ne "")
{
 SendEmail
}

#Open output file after execution
if(!($HideSummaryAtEnd))
{
 if((Test-Path -Path $ExportCSV) -eq "True")
 {   
  Write-Host `n " Exported report has $Count signin activities" 
  Write-Host `n "The Output file availble in:" -ForegroundColor Yellow `n
  Write-Host   CSV format: "$ExportCSV" `n 
  Write-Host   HTML format: $ExportHTML
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 3500+ Microsoft 365 reports and 450+ management actions. ~~" -ForegroundColor Green `n`n
  $Prompt = New-Object -ComObject wscript.shell
  $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
  if ($UserInput -eq 6)
  {
    Invoke-Item "$ExportCSV"
    Invoke-Item "$ExportHTML"
  }   
 }
 else
 {
  Write-Host "No logs found" 
 }
}