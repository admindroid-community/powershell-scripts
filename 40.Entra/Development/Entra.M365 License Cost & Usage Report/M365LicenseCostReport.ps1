<#
=============================================================================================

Name         : Export Microsoft 365 License Cost & Usage Report Using PowerShell  
Version      : 1.1
website      : o365reports.com

-----------------
Script Highlights
-----------------
1. This script allows you to generate nicely formatted 2 CSV files of users’ license cost report and license usage & cost report in the organization.   
2. Helps to generate license cost report for inactive users. 
3. Results can be filtered to lists cost spent on never logged in users only.   
4. Exports disabled users’ license costs alone.  
5. Exports the cost of licenses for external users exclusively.  
6. Identify the Overlapping Licenses assigned to users. 
7. The script uses MS Graph PowerShell and installs MS Graph PowerShell SDK (if not installed already) upon your confirmation.  
8. The script can be executed with an MFA enabled account too. 
9. The script is schedular-friendly. 
10. It can be executed with certificate-based authentication (CBA) too.  

For detailed Script execution:  https://o365reports.com/2024/06/12/export-microsoft-365-license-cost-report-using-powershell/
============================================================================================
\#>

param (
    [string] $CertificateThumbprint,
    [string] $AppId,
    [string] $TenantId,
    [string] $UserCsvPath,
    [string] $Currency,
    [int] $InactiveDays,
    [switch] $NeverLoggedInUsersOnly,
    [switch] $ExternalUsersOnly,
    [switch] $EnabledUsersOnly,
    [switch] $DisabledUsersOnly,
    [switch] $LicenseOverlapingUsersOnly
)

#Function to check MgGraph module and connect to MgGraph.
function ConnectMgGraph{
    #Check MgGraph Beta Module and Install.
    $MsGraphBetaModule =  Get-Module Microsoft.Graph.Beta -ListAvailable

    if($MsGraphBetaModule -eq $null)
    { 
        Write-Host "Important: Microsoft Graph Beta module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $Confirm = Read-Host Are you sure you want to install Microsoft Graph Beta module? [Y] Yes [N] No  
        
        if($Confirm -match "[yY]") 
        { 
            Write-host "Installing Microsoft Graph Beta module....."
            Install-Module Microsoft.Graph.Beta -Scope CurrentUser -AllowClobber 
            Write-host "Microsoft Graph Beta module is installed in the machine successfully." -ForegroundColor Magenta 
        } 
        
        else
        { 
            Write-Host "Exiting.`nNote: Microsoft Graph Beta module must be available in your system to run the script." -ForegroundColor Red
            Exit 
        } 
    }

    #Disconnect MgGraph if already connected.
    if( (Get-MgContext) -ne $null )
    {
        Disconnect-MgGraph | Out-Null
    }

    Write-Host Connecting to Microsoft Graph...

    #Connect to MgGraph via certificate.
    if (($CertificateThumbprint -ne "") -and ($AppId -ne "") -and ($TenantId -ne "")) 
    {
        Connect-MgGraph -TenantId $TenantId -AppId $AppId -CertificateThumbprint $CertificateThumbprint -NoWelcome
        if( (Get-MgContext) -ne $null )
        {
            Write-Host "Connected to Microsoft Graph PowerShell using "(Get-MgContext).AppName" Application Certificate." -ForegroundColor Green
        }
    }

    #Connect to MgGraph.
    else
    {
        Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Directory.Read.All" -NoWelcome
        if( (Get-MgContext) -ne $null )
        {
            Write-Host "Connected to Microsoft Graph PowerShell using" (Get-MgContext).Account "account.`n" -ForegroundColor Green
        }
    }

    #Check connection error.
    if ( ($ConnectionError -ne $null) -or ((Get-MgContext) -eq $null) )
    {    
        Exit
    }
}

#Function for the license usage summary report .
function LicenceUsageReport
{
    Write-Host "Fetching Licenses....`n"
    $Count =0
    $TotalConsumedUnitsCost=0
    $TotalPurchasedUnitsCost=0
    $TotalUnusedUnitsCost=0

    #result path for Organization license report.
    $Location=Get-Location
    $Global:organizationLicenseResultPath= "$Location\LicenseUsageReport "+$DateTime+".csv"

    #Get all the license used by the organization.
    Get-MgBetaSubscribedSku | Select-Object  SkuId , ConsumedUnits , @{Name="PurchasedUnits"; Expression={$_.PrepaidUnits.Enabled} } |
    ForEach-Object {
        $Count++
        $SkuId = $_.'SkuId'
        $ProductDisplayName = $SkuIdDictionary.$SkuId[0]
        $Cost = $SkuIdDictionary.$SkuId[1]
        Write-Progress -Activity "Processing `"$ProductDisplayName`" Subscription " -Status "Processed Subscription Count: $Count"


        #Get cost of license which is unknown.   
        if($Cost -eq '_')
        {     
            Write-host "Enter The Cost for" -NoNewline
            Write-Host " $ProductDisplayName " -ForegroundColor Magenta -NoNewline
            write-Host "License :" -NoNewline
            $Cost = Read-Host  
            $SkuIdDictionary.$SkuId[1] = $Cost ;
        }

        #Calculation.
        $Cost = [decimal] $Cost   
        $UnusedUnits = ([int]$_.'PurchasedUnits'- [int]$_.'ConsumedUnits' )
        [decimal]$ConsumedUnitsCost =($_.'ConsumedUnits' * $Cost )
        [decimal]$PurchasedUnitsCost = ($_.'PurchasedUnits' * $Cost )
        [decimal]$UnusedUnitsCost = $UnusedUnits * $Cost
        [decimal]$TotalConsumedUnitsCost+=$ConsumedUnitsCost
        [decimal]$TotalPurchasedUnitsCost+=$PurchasedUnitsCost
        [decimal]$TotalUnusedUnitsCost+=$UnusedUnitsCost

        #Export into CSV file. 
        $OrganizationLicenseDetail = @{'License Name' = $ProductDisplayName;'Cost'=$Currency+$Cost;'Consumed Units'= $_.'ConsumedUnits'; 'Purchased Units'=$_.'PurchasedUnits' ; 'Unused Units'=$UnusedUnits ; 'Consumed Units Cost'=$Currency+$ConsumedUnitsCost;'Purchased Units Cost'=$Currency+$PurchasedUnitsCost ; 'Unused Units Cost'=$Currency+$UnusedUnitsCost ; 'SkuID'=$skuID}
        
        $OrganizationLicenseDetailObject = New-Object PSObject -Property $OrganizationLicenseDetail
        $OrganizationLicenseDetailObject | Select-object 'License Name','Cost','Purchased Units','Consumed Units','Unused Units','Purchased Units Cost','Consumed Units Cost','Unused Units Cost','SkuID' | Export-csv -path $Global:organizationLicenseResultPath  -NoType -Append -Force
    
    }
    #Add New Line to differentiate the total.
    $NewLine=""
    $NewLine | Add-Content -Path $Global:organizationLicenseResultPath

    #Export the total cost for the license.
    $OrganizationLicenseTotalCost = @{'License Name'="Total"; 'Cost'='-' ;'Purchased Units'= '-' ;'Consumed Units' = '-';'Unused Units' ='-';'Consumed Units Cost'= $Currency+$TotalConsumedUnitsCost;'Purchased Units Cost'=$Currency+$TotalPurchasedUnitsCost ;'Unused Units Cost'=$Currency+$TotalUnusedUnitsCost ; 'SkuID' = '-'}
    
    $OrganizationLicenseTotalCostObject = New-Object PSObject -Property $OrganizationLicenseTotalCost
    $OrganizationLicenseTotalCostObject| Select-object 'License Name','Cost','Purchased Units','Consumed Units','Unused Units','Purchased Units Cost','Consumed Units Cost','Unused Units Cost','SkuID' | Export-csv -path $Global:organizationLicenseResultPath  -NoType -Append -Force
    

}

#Funtion to process the data and export.
function LicensedUserExport
{
    param(
        [Array]  $AssignedLicenses,
        [string] $UserPrincipalName,
        [object] $User
    )

    $Global:UserLicenseResultPath= "$Location\UsersLicenseCostReport "+$DateTime+".csv"

    #SignInDateTime
    $LastSignInDateTime=if($User.SignInActivity.LastSignInDateTime)
                            {$User.SignInActivity.LastSignInDateTime}
                        else
                            {'-'}

    $UserNoInActiveDays=if($User.SignInActivity.LastSignInDateTime)
                            {(New-TimeSpan -Start $LastSignInDateTime).Days}
                        else
                            {"Never Logged In"}

    $LastSuccessfulSignInDateTime=if($User.SignInActivity.LastSuccessfulSignInDateTime)
                                        {$User.SignInActivity.LastSuccessfulSignInDateTime}
                                  else
                                        {'-'}

    if($InactiveDays)
    {
        if($UserNoInActiveDays -eq "Never Logged In" )
        {
            Return
        }

        elseif($InactiveDays -gt $UserNoInActiveDays)
        {
            Return
        }
    }

    elseif(($NeverLoggedInUsersOnly) -and ($UserNoInActiveDays -ne "Never Logged In" ))
    {
        Return
    }

    #UserType
    $UserType = if($User.ExternalUserState -ne $null )
                    {"External"}
                elseif($User.UserType -eq "Guest")
                    {"Interal Guest"}
                else
                    {"Intenal"}

    if($ExternalUsersOnly -and ($UserType -eq "Intenal") )
    {
        Return
    }

    #AccountStatus
    $AccountStatus=if($User.AccountEnabled -eq $true)
                        {"Enabled"} 
                   else 
                        {"Disabled"}

    if($EnabledUsersOnly -and ($AccountStatus -ne "Enabled") )
    {
        Return
    }

    if($DisabledUsersOnly -and ($AccountStatus -ne "Disabled"))
    {
        Return

    }


    $UsageLocation=if($User.UsageLocation)
                        {$User.UsageLocation} 
                   else
                        {'-'}

    $CreatedDateTime=$User.CreatedDateTime

    $JobTitle=if($User.JobTitle)
                    {$User.JobTitle} 
              else
                    {'-'}

    $Department=if($User.Department)
                    {$User.Department} 
                else
                    {'-'}

    $Cost=[decimal]0
    $ProductDisplayNameArray=@()
    $SkuIdArray=@()

    if($AssignedLicenses)
    {
        $AssignedLicenses | 
        ForEach-Object {
            $SkuIdNow=$_.SkuId
            $SkuIdArray+=$SkuIdNow
            $ProductDisplayNameArray += $SkuIdDictionary.$SkuIdNow[0]
            $Cost += [decimal]$SkuIdDictionary.$SkuIdNow[1]
        }

        $ProductDisplayName= $ProductDisplayNameArray -join ', '
        $SkuId= $skuIdArray -join ', '
        $DirectlyAssignedArray=@()
        $GroupBasedAssignedArray=@()
        $GroupBasedwithGroupArray=@()
        $InBothDirectAndGroupArray=@()
        $NoOfDuplicateArray=@()

        $user.licenseAssignmentStates | 
        ForEach-Object{
            if($_.AssignedByGroup)
            {
                $GroupBasedAssignedArray += $SkuIdDictionary.($_.SkuId)[0]

                if($GroupIdDictinary.($_.AssignedByGroup))
                {
                    $GroupName = $GroupIdDictinary.($_.AssignedByGroup)
                }

                else
                {
                    $GroupName = Get-MgBetaGroup -GroupId $_.AssignedByGroup | Select-Object -ExpandProperty DisplayName 
                    $GroupIdDictinary[$_.AssignedByGroup] = $GroupName
                }

                $GroupBasedwithGroupArray += $SkuIdDictionary.($_.SkuId)[0] + " [$GroupName]"
            }

            else
            {
                $DirectlyAssignedArray += $SkuIdDictionary.($_.SkuId)[0]                 
            }
        }

        $InBothDirectAndGroupArray= $GroupBasedAssignedArray | Where-Object { $DirectlyAssignedArray -contains $_}
        $NoOfDuplicateArray= $InBothDirectAndGroupArray | Group-Object | 
        ForEach-Object {
            "$($_.Name) [$($_.Count+1)]"
        }

        #handling if duplicate assigned only through Groupbased assigning.
        $NoOfDuplicateArray += (($GroupBasedAssignedArray | Where-Object { $DirectlyAssignedArray -notcontains $_})| Group-Object) | 
        ForEach-Object {
            if($_.Count -gt 1)
                {
                    "$($_.Name) [$($_.Count)]"
                }
        }

        $InBothDirectAndGroupArray +=(($GroupBasedAssignedArray | Where-Object { $DirectlyAssignedArray -notcontains $_})| Group-Object) | 
        ForEach-Object {
            if($_.Count -gt 1)
                {
                    "$($_.Name)"
                }
        }

        #Check the value if null convert to '-'.

        $NoOfDuplicate=if($NoOfDuplicateArray)
                            {$NoOfDuplicateArray -join ', '} 
                       else 
                            {'-'}

        $IsDuplicateLicense = if($InBothDirectAndGroupArray)
                                    {"Yes"} 
                               else
                                    {"No"}

        $DirectlyAssigned= if($DirectlyAssignedArray)
                                {$DirectlyAssignedArray -join ', '} 
                           else
                                {'-'}

        $GroupBasedAssigned = if($GroupBasedwithGroupArray)
                                    {$GroupBasedwithGroupArray -join ', '} 
                              else
                                    {'-'}

    }

    else
    {
        $ProductDisplayName='-'
        $SkuId='-'
        $InBothDirectAndGroup="-"
        $DirectlyAssigned="-"
        $GroupBasedAssigned="-"
        $NoOfDuplicate="-"
    }

    #Filter the LicenseOverlapingUsersOnly
    if($LicenseOverlapingUsersOnly)
    {
        if($IsDuplicateLicense -ne "Yes")
        {
            Return
        }  
    }
    $Global:Count++
    Write-Progress -Activity "Processing User `"$DisplayName`"" -Status "Processed User Count: $Global:Count"

    #Export into CSV file.
    $UserLicenseDetail = @{'Display Name' = $DisplayName;'User Principal Name'=$UserPrincipalName;'Cost'=$Currency+$Cost;'Assigned Licenses'=$ProductDisplayName; 'SkuID'=$SkuId ; 'Directly Assigned Licenses'=$DirectlyAssigned ; 'Licenses Assigned Via Groups'=$GroupBasedAssigned ;  'Is Duplicate License' = $IsDuplicateLicense;'Dublicate Licenses Found With Count'= $NoOfDuplicate; 'Account Status' = $AccountStatus  ; 'Usage Location' = $UsageLocation ;'Created Date Time'=$CreatedDateTime ;'Job Title' = $JobTitle ; 'Department' =$Department ;'Last Sign In Date Time'=$LastSignInDateTime ; 'Inactive Days'=$UserNoInActiveDays;'Last Successful SignIn DateTime'=$LastSuccessfulSignInDateTime ; 'User Type'=$UserType }
    
    $UserLicenseDetailObject = New-Object PSObject -Property $UserLicenseDetail
    $UserLicenseDetailObject| Select-object 'Display Name','User Principal Name','Assigned Licenses','Cost','Directly Assigned Licenses','Licenses Assigned Via Groups','Is Duplicate License','Dublicate Licenses Found With Count','Account Status','Job Title','Department','Created Date Time' ,'Last Sign In Date Time' ,'Inactive Days','Last Successful SignIn DateTime','User Type','Usage Location' | Export-csv -path $Global:UserLicenseResultPath  -NoType -Append -Force
}

#Function to Licensed Users.
function AllLicensedUserReport
{
    Write-Host "Generating license cost report...."

    $Global:Count = 0
    #Get all Licensed users.
    Get-MgBetaUser -All -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable Records -Property DisplayName, UserPrincipalName, AssignedLicenses , LicenseAssignmentStates, AccountEnabled, UsageLocation, CreatedDateTime, JobTitle, Department , UserType , SignInActivity , ExternalUserState | 
    ForEach-Object { 
        $DisplayName=$_.DisplayName
        $UserPrincipalName=$_.UserPrincipalName
        LicensedUserExport -AssignedLicenses $_.AssignedLicenses -UserPrincipalName $UserPrincipalName -User $_
    }


}

#Function to Selected users.
function SelectedUserReport
{
    #Check the CSV path.
    if (-not (Test-Path $UserCsvPath –PathType Leaf)) 
    {
        Write-Host "`n$UserCsvPath does`'t contain CSV file." -ForegroundColor Red
        Write-Host "Enter a valid path and try again." -ForegroundColor Yellow
        Return
    }

    Write-Host "Generating license cost report....`n"
    $Global:Count = 0

    #Import the UserId from the given CSV file. 
    Import-Csv $UserCsvPath | Select-Object UserId -Unique |
    ForEach-Object {
        Get-MgBetaUser -UserId $_.UserId  -Property DisplayName, UserPrincipalName, AssignedLicenses , LicenseAssignmentStates, AccountEnabled, UsageLocation, CreatedDateTime, JobTitle, Department , UserType , SignInActivity , ExternalUserState  | 
        ForEach-Object { 
            $DisplayName=$_.DisplayName
            $UserPrincipalName=$_.UserPrincipalName
            LicensedUserExport -AssignedLicenses $_.AssignedLicenses -UserPrincipalName $UserPrincipalName -User $_
        }
    }

  
}

#--------------------------------------------------------------------------------Main Function Starts--------------------------------------------------------------------------------#

#Connecting to MgGraph.
ConnectMgGraph

$DateTime=(Get-Date -Format "dd-MM-yy hh_mm_ss tt").ToString()

#Convert the Data CSV into Hash table.
$SkuIdDictionary = @{}

Import-Csv "$PSScriptRoot\LicenseCostAndUserFriendlyName.csv" | 
ForEach-Object {
    $SkuId = $_.'SkuId'
    $Name = $_.'License Name'
    $Cost = $_.'Cost'
    $SkuIdDictionary[$SkuId] = @($Name, $Cost)
}

$GroupIdDictinary=@{}

#Calling LicenseUsageReport Function.
LicenceUsageReport

#Calling SelectedUserReport Function.
if($UserCsvPath )
{
    SelectedUserReport
}

#Calling AllLicensedUserReport Function.
else
{
    AllLicensedUserReport
}

$Prompt = New-Object -ComObject wscript.shell   
#Check The $Global:UserLicenseResultPath path and open file.
if(((Test-Path $Global:UserLicenseResultPath)  -eq "True") -and ((Test-Path $Global:organizationLicenseResultPath) -eq "True"))
{   
    Write-Host "`nDetailed license usage and users' license cost reports are stored in: $Location" -ForegroundColor Cyan
    $UserInput = $Prompt.popup("Do you want to open output files?",` 0,"Open Output Files",4)   
    If ($UserInput -eq 6)
    {   
        Invoke-Item "$Global:organizationLicenseResultPath" ,"$Global:UserLicenseResultPath"
    }
    
}

elseif ((Test-Path $Global:organizationLicenseResultPath) -eq "True")
{
    Write-Host "`nDetailed license usage & cost report stored in: $Location" -ForegroundColor Cyan
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)   
    If ($UserInput -eq 6)
    {   
        Invoke-Item "$Global:organizationLicenseResultPath"
    }
    
}
else
{
 Write-Host "There are no users matching the specified filters"
}

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    

Disconnect-MgGraph | Out-Null

#--------------------------------------------------------------------------------Main Function Ends--------------------------------------------------------------------------------#
