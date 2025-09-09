<#
=============================================================================================
Name:           Get Nested Groups in Microsoft 365 Using PowerShell
Version:        1.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~~

1. The script exports 2 different CSV reports: 
   i) M365 nested groups summary report 
   ii) M365 nested groups detailed report 
2. Exports nested group reports based on group types, such as Security, Distribution, Mail enabled security group. 
3. Automatically installs the required Microsoft Graph module with your confirmation. 
4. The script can be executed with an MFA-enabled account too. 
5. Supports Certificate-based Authentication (CBA) too. 
6. The script is scheduler friendly. 

For detailed script execution:  https://o365reports.com/2024/11/19/get-nested-groups-in-microsoft-365-using-powershell/
============================================================================================
#>

Param 
( 
    [Parameter(Mandatory = $false)] 
    [switch]$DistributionList, 
    [switch]$Security, 
    [switch]$MailEnabledSecurity, 
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
) 

$SummaryReport ="$(Get-Location)\M365NestedGroups_SummaryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$DetailedReport ="$(Get-Location)\M365NestedGroups_DetailedReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 


# Function to connect to Microsoft Graph
function Connect_ToMgGraph {
    # Check if Microsoft Graph module is installed
    $MsGraphModule = Get-Module Microsoft.Graph -ListAvailable
    if ($MsGraphModule -eq $null) {
        Write-Host "`nImportant: Microsoft Graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host "Are you sure you want to install Microsoft Graph module? [Y] Yes [N] No"
        if ($confirm -match "[yY]") {
            Write-Host "Installing Microsoft Graph module..."
            Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber
            Write-Host "Microsoft Graph module is installed in the machine successfully" -ForegroundColor Magenta 
        } else {
            Write-Host "Exiting. `nNote: Microsoft Graph module must be available in your system to run the script" -ForegroundColor Red
            Exit
        }
    } 

    Write-Host "`nConnecting to Microsoft Graph..."
    
    if (($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne "")) {
        # Use certificate-based authentication if TenantId, ClientId, and CertificateThumbprint are provided
        Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    } else {
        # Use delegated permissions (Scopes) if credentials are not provided
        Connect-MgGraph -Scopes "GroupMember.Read.All" -NoWelcome 
    }

    # Verify connection
    if ((Get-MgContext) -ne $null) {
        Write-Host "Connected to Microsoft Graph PowerShell using account: $((Get-MgContext).Account)" -ForegroundColor Yellow
    } else {
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
}

# Function to get report details
Function Get_members {
    $ParentGroup = $_.DisplayName
    $ParentGroupEmailAddress = $_.Mail
    $ParentGroupId = $_.Id
    $CreatedDateTime = $_.CreatedDateTime
    
    $TransitiveNestedGroups = Get-MgGroupTransitiveMemberAsGroup -All -GroupId $ParentGroupId 
    $TransitiveNestedGroupsCount = $TransitiveNestedGroups.Count
    $TransitiveNestedGroupNames = $TransitiveNestedGroups.DisplayName -join ", "

    Write-Progress -Activity "`n     Processed group count: $Count "`n"  Checking members of: $ParentGroup"

    # Check for Empty Group
    if($TransitiveNestedGroupsCount -eq 0) { return }
    
    # Get the parent group type based on properties
    if($_.Mail -ne $null) {
        if($_.SecurityEnabled -eq $false) {
            $ParentGroupType="DistributionList"
        }
        else {
            $ParentGroupType="MailEnabledSecurity"
        }
    }
    else {
        $ParentGroupType="Security"
    }

    # If parent group email address is empty
    if($ParentGroupEmailAddress -eq $null) {
        $ParentGroupEmailAddress = "-"
    }

    # Filter for security group
    if(($Security.IsPresent) -and ($ParentGroupType -ne "Security")) {
        return
    }

    # Filter for Distribution list
    if(($DistributionList.IsPresent) -and ($ParentGroupType -ne "DistributionList")) {
        return
    }

    # Filter for mail enabled security group
    if(($MailEnabledSecurity.IsPresent) -and ($ParentGroupType -ne "MailEnabledSecurity")) {
        return
    }

    Print_SummaryReportContent

    foreach($TransitiveNestedGroup in $TransitiveNestedGroups) {
        $NestedGroup = $TransitiveNestedGroup.displayName
        $NestedGroupEmailAddress = $TransitiveNestedGroup.mail
        $NestedGroupId = $TransitiveNestedGroup.Id
        $NestedUsersCount = Get-MgGroupTransitiveMemberCountAsUser -GroupId $NestedGroupId -ConsistencyLevel eventual
        
        # If Nested group email address is empty
        if($NestedGroupEmailAddress -eq $null) {
            $NestedGroupEmailAddress = "-"
        }

        # Get the Nested group type based on properties
        if($TransitiveNestedGroup.Mail -ne $null) {
            if($TransitiveNestedGroup.SecurityEnabled -eq $false) {
                $NestedGroupType="DistributionList"
            }
            else {
                $NestedGroupType="MailEnabledSecurity"
            }
        }
        else {
            $NestedGroupType="Security"
        }
            
        Print_DetailedReportContent
    }
}


# Print Summary Report
Function Print_SummaryReportContent {
    $Result=@{'Group Name'=$ParentGroup; 'Group Type'=$ParentGroupType; 'Group Email Address'=$ParentGroupEmailAddress; 'Nested Group Names'=$TransitiveNestedGroupNames; 'Nested Groups Count'=$TransitiveNestedGroupsCount; 'Group Id'=$ParentGroupId;} 
    $Results= New-Object PSObject -Property $Result 
    $Results | Select-Object 'Group Name', 'Group Type', 'Group Email Address', 'Nested Group Names', 'Nested Groups Count', 'Group Id' | Export-Csv -Path $SummaryReport -Notype -Append
}

# Print Detailed Output
Function Print_DetailedReportContent {
    $Result=@{'Parent Group Name'=$ParentGroup; 'Nested Group Name'=$NestedGroup; 'Nested Group Type'=$NestedGroupType; 'Nested Group Email Address'=$NestedGroupEmailAddress; 'Members Count in Nested Group'=$NestedUsersCount; }  
    $Results= New-Object PSObject -Property $Result 
    $Results | Select-Object 'Parent Group Name', 'Nested Group Name', 'Nested Group Type', 'Nested Group Email Address', 'Members Count in Nested Group' | Export-Csv -Path $DetailedReport -Notype -Append
}

# Connecting to the Microsoft Graph PowerShell Module
Connect_ToMgGraph

# Get all nested groups in Microsoft 365
$Count = 0
Get-MgGroup -All | ForEach-Object {
    $Count++
    Get_Members
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

# Open output file after execution 
if((Test-Path -Path $DetailedReport) -eq "True" -and (Test-Path -Path $SummaryReport) -eq "True") {
    Write-Host `n" Nested groups summary report available in: " -NoNewline -ForegroundColor Yellow
	Write-Host $SummaryReport
    Write-Host `n" Nested groups detailed report available in: " -NoNewline -ForegroundColor Yellow
	Write-Host $DetailedReport
	Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1900+ Microsoft 365 reports. ~~" -ForegroundColor Green
    $Prompt = New-Object -ComObject wscript.shell  
    $UserInput = $Prompt.popup("Do you want to open output files?",` 0,"Open Output File",4)  
    If ($UserInput -eq 6) {  
        Invoke-Item "$DetailedReport"  
        Invoke-Item "$SummaryReport"  
    } 
}
else {
    Write-Host `n"No nested groups are found for the specified criteria." -ForegroundColor Red
}