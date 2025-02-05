<#
=============================================================================================
Name:           Find All Team Channels with External Members in Microsoft 365   
Version:        1.0
Website:        o365reports.com

Script Highlights:  
~~~~~~~~~~~~~~~~~
1. The script exports output into 2 CSV files: 
     -Detailed report: Teams Channels and their external user details
     -Summary report: Teams Channels and their guest users count 
2. Exports external users across all Teams channel types, such as private, shared, and standard.
3. Exports all team channels an external user is a member of.
4. Automatically install the Microsoft Graph PowerShell module (if not installed already) upon your confirmation.
5. The script can be executed with an MFA-enabled account too.
6. It can be executed with certificate-based authentication (CBA) too. 
7. The script is schedular-friendly.

For detailed Script execution:  https://o365reports.com/2025/02/04/get-all-teams-channels-with-external-members/
============================================================================================
#>

Param 
( 
    [Parameter(Mandatory = $false)] 
    [string]$TeamName,
    [string]$ChannelName,
    [string]$MemberUPN,
    [ValidateSet(
        'Standard',
        'Private',
        'Shared'
    )]
    [string]$ChannelType,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
) 

$SummaryReport ="$(Get-Location)\TeamsChannelsWithExternalMembers_SummaryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$DetailedReport ="$(Get-Location)\TeamsChannelsWithExternalMembers_DetailedReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 

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

    Write-Host "Connecting to Microsoft Graph..."
    
    if (($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne "")) {
        # Use certificate-based authentication if TenantId, ClientId, and CertificateThumbprint are provided
        Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome 
    } else {
        # Use delegated permissions (Scopes) if credentials are not provided
        Connect-MgGraph -Scopes "Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMember.Read.All" -NoWelcome 
    }

    $Script:MgContext = Get-MgContext

    # Verify connection
    if (($Script:MgContext) -ne $null) {
        Write-Host "Connected to Microsoft Graph PowerShell using account: $(($Script:MgContext).Account)" 
    } else {
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
}

# Print Summary Report
Function Print_SummaryReportContent {
    $Result = @{'Team Name'=$TeamDisplayName; 'Channel Name'=$ChannelDisplayName; 'Membership Type'=$MembershipType; 'External Members Count'=$ExternalMembersCount; 'External Members'=$ExternalMembers; 'Team Id'=$TeamId; 'Channel Id'=$ChannelId;} 
    $Results = New-Object PSObject -Property $Result 
    $Results | Select-Object 'Team Name', 'Channel Name', 'Membership Type', 'External Members Count', 'External Members', 'Team Id', 'Channel Id' | Export-Csv -Path $SummaryReport -Notype -Append
}

# Print Detailed Output
Function Print_DetailedReportContent {
    $Result = @{'Team Name'=$TeamDisplayName; 'Channel Name'=$ChannelDisplayName; 'Membership Type'=$MembershipType; 'External Member Name'=$ExternalMemberName; 'External Member Email'=$ExternalMemberEmail; 'External Member Id'=$ExternalMemberGuid;} 
    $Results = New-Object PSObject -Property $Result 
    $Results | Select-Object 'Team Name', 'Channel Name', 'Membership Type', 'External Member Name', 'External Member Email', 'External Member Id'| Export-Csv -Path $DetailedReport -Notype -Append
}

# Connecting to the Microsoft Graph PowerShell Module
Connect_ToMgGraph

# Get all teams channels with external members in Microsoft 365
$ProcessedTeamsCount = 0
$TeamsChannelsWithExternalMembersCount = 0
$TenantId = ($Script:MgContext).TenantId

if (!([string]::IsNullOrEmpty($TeamName))) { $TeamFilter = "displayName eq '$($TeamName)'" }

if (!([string]::IsNullOrEmpty($ChannelName))) { 
    if (!([string]::IsNullOrEmpty($TeamName))) {
        $ChannelFilter = "displayName eq '$($ChannelName)'" 
    }
    else {
        Write-Host "`nError: TeamName param is mandatory to filter based on Channels." -ForegroundColor Red
        Exit
    }
}

Get-MgTeam -Filter "$($TeamFilter)" -All | ForEach-Object {
    $TeamId = $_.Id
    $TeamDisplayName = $_.DisplayName
    $ProcessedTeamsCount++

    Get-MgTeamChannel -Filter "$($ChannelFilter)" -All -Team $TeamId | ForEach-Object {
        $ChannelId = $_.Id
        $ChannelDisplayName = $_.DisplayName
        $MembershipType = $_.MembershipType
        $ExternalMembersCount = 0
        $ExternalMembers = @()
        $ExternalUserFound = $false
        Write-Progress -Activity "`n     Processed Teams count: $ProcessedTeamsCount"`n"  Processing Channel: $($ChannelDisplayName) from Team: $($TeamDisplayName)"

        Get-MgTeamChannelMember -All -Team $TeamId -Channel $ChannelId | ForEach-Object {
            $ExternalMemberTenantId  = $_.AdditionalProperties["tenantId"]
            if (($ExternalMemberTenantId -ne $TenantId) -or ($_.Roles -contains "guest")){
                $ExternalMemberName      = $_.DisplayName
                $ExternalMemberEmail     = $_.AdditionalProperties["email"]
                $ExternalMemberGuid      = $_.AdditionalProperties["userId"]

                if (($MembershipType -notin @("standard","private"))) { $MembershipType = "shared" }

                # Apply filters based on params
                if (!([string]::IsNullOrEmpty($MemberUPN)) -and ($MemberUPN -ne $ExternalMemberEmail)) { return }
                if (!([string]::IsNullOrEmpty($ChannelType)) -and ($ChannelType -ne $MembershipType)) { return }

                $ExternalUserFound = $true
                $ExternalMembersCount++
                $ExternalMembers += $ExternalMemberEmail 
                Print_DetailedReportContent
            }
        }
        if ($ExternalUserFound) {
            $ExternalMembers = $ExternalMembers -join ", "
            $TeamsChannelsWithExternalMembersCount++
            Print_SummaryReportContent
        }
    }
}
# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1900+ Microsoft 365 reports. ~~" -ForegroundColor Green

# Open output file after execution 
if((Test-Path -Path $DetailedReport) -eq "True") {
    Write-Host `n"$TeamsChannelsWithExternalMembersCount teams channels are found with external members." 
    Write-Host `n"  The summary report available in: " -NoNewline -ForegroundColor Yellow; Write-Host $SummaryReport
    Write-Host `n"  The detailed report available in: " -NoNewline -ForegroundColor Yellow; Write-Host $DetailedReport
	$Prompt = New-Object -ComObject wscript.shell  
    $UserInput = $Prompt.popup("Do you want to open output files?",` 0,"Open Output File",4)  
    If ($UserInput -eq 6) {  
        Invoke-Item "$SummaryReport"
        Invoke-Item "$DetailedReport"  
    } 
}
else {
    Write-Host `n"No teams channels are found with external members." -ForegroundColor Red
}