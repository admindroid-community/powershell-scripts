<#
=============================================================================================
Name:           Find Unused Licenses in Microsoft 365 Using PowerShell
Description:    This script exports a report of unused Microsoft 365 licenses by identifying inactive users through their last successful sign-in activity.
Version:        1.0
website:        o365reports.com

Script Highlights:
~~~~~~~~~~~~~~~~~

1. Retrieves unused licenses based on users' last successful sign-in time.
2. Lists licenses assigned to sign-in disabled users.
3. Identifies licenses assigned to never logged-in user accounts.
4. Filters unused licenses by type, such as paid, free, or trial.
5. Fetches inactive licenses assigned to external accounts.
6. Identifies unused specific licenses, such as Power BI Pro.
7. Automatically verifies and installs the Microsoft Graph PowerShell Module (if not already installed) upon your confirmation.
8. Supports Certificate-based Authentication (CBA) too.
9. The script is scheduler-friendly.


For detailed script execution: https://o365reports.com/2025/09/02/find-unused-licenses-in-microsoft-365-using-powershell/ 

============================================================================================
#>
Param(
    [int]$InactiveDays,
    [Nullable[int]]$LicenseCount = $null,
    [string]$ImportCSVPath,
    [switch]$ReturnNeverLoggedInUser,
    [ValidateSet("InternalUser", "ExternalUser")]
    [string]$UserType,
    [ValidateSet("EnabledUser", "DisabledUser")]
    [string]$UserState,
    [ValidateSet( "Paid", "Trial", "Free")]
    [string]$LicenseType,
    [string[]]$LicensePlanList,
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

if (-not $InactiveDays -and -not $ReturnNeverLoggedInUser) {
    do {
        $InactiveDays = Read-Host "`nEnter the number of inactive days"
        if ($InactiveDays -notmatch '^\d+$') {
            Write-Host "Please enter a valid number." -ForegroundColor Red
        }
    } while ($InactiveDays -notmatch '^\d+$')
    $InactiveDays = [int]$InactiveDays
}

Function Connect_MgGraph {
    #Check for module installatiion
    $Module = Get-Module -Name microsoft.graph -ListAvailable
    if($Module.Count -eq 0){
        Write-Host "Microsoft Graph PowerShell SDK is not available"  -ForegroundColor yellow 
        $Confirm = Read-Host Are you sure want to install the module? [Y]Yes [N]No
        if($Confirm -match [Yy]){
            Write-Host "Installing Microsoft Graph PowerShell module..."
            Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
        }
        else{
            Write-Host "Microsoft Graph PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph cmdlet." 
            Exit
        }
    }
    #Disconnect Existing MgGraph session
    if ($CreateSession.IsPresent) {
    	Disconnect-MgGraph | Out-Null
    }
     Write-Host "`nConnecting to Microsoft Graph..." 
    if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
    {  
        Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    }
    else{
        Connect-MgGraph -Scopes  "User.Read.All", "Group.Read.All", "Organization.Read.All", "AuditLog.Read.All" -NoWelcome
    }
}

Connect_MgGraph

$Location = Get-Location
$ExportCSV = "$Location\UnusedM365LicensesByLastSignIn_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$ExportResult = ""
$Count = 0
$PrintedUsers = 0

if( $ImportCSVPath -ne "" -and !(Test-Path $ImportCSVPath)){
    Write-Host "File not found: $($ImportCSVPath)" -ForegroundColor Red
    Exit
}

#Get Licenses
$FriendlyNameHash = @{}
Import-Csv -Path ".\LicenseFriendlyName.csv" -ErrorAction Stop | ForEach-Object{
    $FriendlyNameHash[$_.String_Id] = $_.Product_Display_Name
}
$LicenseMap = @{}
if($ImportCSVPath -ne ""){
    $LicenseNames = Import-Csv -Header "SkuPartNumber" -path $ImportCSVPath | ForEach-Object { $_ }
    foreach ($License in $LicenseNames) {
        $SkuPartNumber = $License.SkuPartNumber
        if ($FriendlyNameHash.ContainsKey($SkuPartNumber)) {
            $LicenseMap[$SkuPartNumber] = $FriendlyNameHash[$SkuPartNumber]
        }
    }
}
else{
    $LicenseMap = $FriendlyNameHash
} 

#Get License type details
$LicenseSkuIdMap = @{}
$LifeCycleDateInfo = Get-MgDirectorySubscription -All 
$LifeCycleDateInfo | ForEach-Object{
    $LicenseSkuIdMap[$_.SkuId] = $_.SkuPartNumber
}

#Retrieve Users
Write-Host "`nRetrieving inactive users with assigned licenses..."
$RequiredProperties = @('UserPrincipalName','DisplayName','SignInActivity','UserType','CreatedDateTime','AccountEnabled', 'LicenseAssignmentStates', 'Department','JobTitle')  
Get-MgUser -All -Property $RequiredProperties | select $RequiredProperties | ForEach-Object{
    $Count++
    $UPN = $_.UserPrincipalName
    Write-Progress -Activity "        Processing  user: $($count) $($UPN)" 
    $DisplayName = $_.DisplayName
    $UserCategory = $_.UserType
    $LastSuccessfulSigninDate = $_.SignInActivity.LastSuccessfulSignInDateTime
    $LastInteractiveSignIn = $_.SignInActivity.LastSignInDateTime
    $LastNon_InterativeSignIn = $_.SignInActivity.LastNonInteractiveSignInDateTime
    $CreatedDate = $_.CreatedDateTime
    $AccountEnabled = $_.AccountEnabled
    $Department = if($_.Department -eq $null) {" -"} else{$_.Department}
    $JobTitle = if($_.JobTitle -eq $null) {" -"} else{$_.JobTitle}
    $TotalLicenses = 0
    $LicenseStates = $_.LicenseAssignmentStates 
    $Print = 1
    
    #Calculate Inactive users days
    if($LastSuccessfulSigninDate -eq $null){
        $LastSuccessfulSigninDate = "Never Logged In"
        $InactiveUserDays = "-"
    }else{
        $InactiveUserDays = (New-TimeSpan -Start $LastSuccessfulSigninDate).Days
    }

    if($LastInteractiveSignIn -eq $null){
        $LastInteractiveSignIn = "Never Logged In"
    }

    if($LastNon_InterativeSignIn -eq $null){
        $LastNon_InterativeSignIn = "Never Logged In"
    }

    #Get account status
    if($AccountEnabled -eq $true){
        $AccountStats = "Enabled"
    }
    else{
        $AccountStats = "Disabled"
    }
    
    #Inactive days based on last successful signins filter
    if ($ReturnNeverLoggedInUser.IsPresent -and ($LastInteractiveSignIn -ne "Never Logged In" -or $LastNon_InterativeSignIn -ne "Never Logged In")) {
        $Print = 0
    }
    elseif (-not $ReturnNeverLoggedInUser.IsPresent) {
        if ($LastSuccessfulSigninDate -eq "Never Logged In") {
            $Print = 0
        }
        # Filter by inactive days
        elseif (($InactiveDays -ne 0) -and ($InactiveDays -ge $InactiveUserDays)) {
            $Print = 0
        }
    }
    
    #Filter for internal users only
    if(($UserType -eq "InternalUser") -and ($UserCategory -eq "Guest")){
        $Print = 0
    }

    #Filter for external users only
    if(($UserType -eq "ExternalUser") -and ($UserCategory -ne "Guest")){
        $Print = 0
    }

    #Signin allowed Users
    if(($UserState -eq "EnabledUser") -and ($AccountStats -eq 'Disabled')){
        $Print = 0
    }

    #Signin disabled Users
    if(($UserState -eq "DisabledUser") -and ($AccountStats -eq 'Enabled')){
        $Print = 0
    }
    
    #Licensed users only
    $LicensePartNumbers = @()
    $Groups = @()
    $GroupLicense = @()
    $DirectLicense = @()
    if($LicenseStates.Count -ne 0){
        foreach($State in $LicenseStates){
            if($State){
                $Flag = 1
                $LicensePartNumber = ""
                $LicenseName = ""
                if($LicenseSkuIdMap.ContainsKey($State.SkuId)){
                    $LicensePartNumber = $LicenseSkuIdMap[$State.SkuId]
                    $LicenseName = $LicenseMap[$LicensePartNumber]
                    $MoreSkuDetails = $LifeCycleDateInfo | Where-Object {$_.skuId -eq $State.SkuId}
                    $ExpiryDate = $MoreSkuDetails.nextLifeCycleDateTime
                    #Filter SkuPartNumber
                    if($LicensePlanList){
                        if($LicensePlanList -notcontains $LicensePartNumber){
                            $Flag = 0
                        }
                    }
                    
                    #Filter Free Licensed User
                    if($LicenseType -eq "Free"){
                        if($ExpiryDate -ne $null){
                            $Flag = 0
                        }
                    }

                    #Filter Trial Licensed User
                    if($LicenseType -eq "Trial"){
                        if(-not $MoreSkuDetails.isTrial){
                            $Flag = 0
                        }
                    }

                    #Filter Paid Licensed User
                    if($LicenseType -eq "Paid"){
                        if(($ExpiryDate -eq $null) -or ($MoreSkuDetails.isTrial)){
                            $Flag = 0
                        }
                    }

                    if($Flag -eq 1){
                        if($LicenseName){
                            $LicensePartNumbers += $LicensePartNumber
                            if($State.AssignedByGroup -ne $null){
                                $Groups += (Get-MgGroup -GroupId $State.AssignedByGroup -ErrorAction SilentlyContinue).DisplayName
                                $GroupLicense += $LicenseName
                            }
                            else{
                                $DirectLicense += $LicenseName
                            }
                        }
                    }
                }
            }
        }

        if(($DirectLicense.Count -ne 0) -or ($GroupLicense.Count -ne 0)) {
            $LicensePlans = $LicensePartNumbers -join ", "
            $TotalLicenses = $DirectLicense.Count + $GroupLicense.Count
            $GroupNames = if($Groups.Count -ne 0) {$Groups -join ","} else {"- "}
            $GroupLicenseNames = if($GroupLicense.Count -ne 0) { $GroupLicense -join ","} else {"- "}
            $DirectLicenseNames = if($DirectLicense.Count -ne 0) { $DirectLicense -join ","} else{"- "}
        }
        else{
            $Print = 0
        }
    }
    else{
        $Print = 0;
    }
    
    #LicenseCount above users only
    if($LicenseCount -ne $null){
        if($LicenseCount -gt $TotalLicenses){
            $Print = 0
        }
    }

    #Export users to output file
    if($Print -eq 1 ){
        $PrintedUsers++
        $ExportResult = [PSCustomObject]@{ 'Display Name' = $DisplayName; 'UPN' = $UPN; 'User Type' = $UserCategory; 'Account Status' = $AccountStats; 'License Plans' = $LicensePlans; 'Directly Assigned Licenses' = $DirectLicenseNames; 'Licenses Assigned via Groups' =$GroupLicenseNames; 'Assigned via (Group Names)' = $GroupNames; 'License Count' = $TotalLicenses;'Last Successful SignIn Date '= $LastSuccessfulSigninDate; 'Inactive Days' = $InactiveUserDays; 'Last Interactive SignIn Date' = $LastInteractiveSignIn; 'Last Non-Interactive SignIn Date' = $LastNon_InterativeSignIn;'Creation Date' = $CreatedDate; 'Department' = $Department;'Job Title' = $JobTitle;}
        $ExportResult | Export-Csv -Path $ExportCSV -NoTypeInformation -Append
    }
}

Disconnect-MgGraph | Out-Null

Write-Host `nScript executed successfully.
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green `n`n

#Open output file after execution
if(((Test-Path -Path $ExportCSV) -eq "True"))
{
    Write-Host "Exported report has $($PrintedUsers) user(s)." 
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)
    if ($UserInput -eq 6)
    {  
        Invoke-Item "$ExportCSV"
    }
    Write-Host "The generated report is available in:" -NoNewline -ForegroundColor Yellow; Write-Host "$($ExportCSV)"
}
else
{
    Write-Host "No user found" -ForegroundColor Red
}