 <#
=============================================================================================
Name: Manage and Report Microsoft Teams Tag Members Using PowerShell
Version: 1.0
Website: o365reports.com

~~~~~~~~~~~~~~~~~
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script uses MS Graph PowerShell and installs MS Graph PowerShell SDK (if not installed already) upon your confirmation.  
2. Exports tags for all teams in your entire tenant. 
3. Helps you to add user(s) to a tag. 
4. Allows you remove user(s) from a tag.   
5. The script is schedular-friendly.  
6. Exports report results as a CSV file. 

For detailed script execution: https://o365reports.com/2024/07/03/manage-and-report-microsoft-teams-tag-members-using-powershell/

=============================================================================================
#> 
param(
    [string]$TenantID,
    [string]$ClientID,
    [string]$CertificateThumbPrint,
    [string]$TeamId,
    [string]$TeamworkTagId,
    [string]$UserId,
    [string]$UserTeamworkTagMemberId,
    [string]$CSV,
    [switch]$GetTeamTagReports,
    [switch]$AddUsertoTag,
    [switch]$RemoveUserFromTag

)
$logfile = ".\LogFile_" + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".txt"
Function WriteToLogFile ($message) {
    $message >> $logfile
}
function InvokeLogFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExportCSV
    )
    Write-Host `n Script executed successfully -ForegroundColor Green
    if ((Test-Path -Path $ExportCSV) -eq "True") {

        Write-Host `n "The Log file availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$ExportCSV" `n
        $Prompt = New-Object -ComObject wscript.shell  
        $UserInput = $Prompt.popup("Do you want to open log file?", 0, "Open Output File", 4)  
        if ($UserInput -eq 6) {  
            Invoke-Item "$ExportCSV"  
        } 
    }
}

function Get-UserInput {
    param(
        [string] $Prompt
    )
    Write-Host $Prompt -ForegroundColor Yellow
    do {
        $value = Read-Host 
        $value = $value.Trim()
    }until(-not [string]::IsNullOrEmpty($value))
    return ($value)
}
function InvokeCSVFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExportCSV
    )
    Write-Host `n Script executed successfully -ForegroundColor Green
    if ((Test-Path -Path $ExportCSV) -eq "True") {
   
        Write-Host `n "The Output file availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$ExportCSV" `n 
        $Prompt = New-Object -ComObject wscript.shell  
        $UserInput = $Prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)  
        if ($UserInput -eq 6) {  
            Invoke-Item "$ExportCSV"  
        } 
    }
}


function Get-TeamTagReports {
    try {
        $FilePath = ".\TenantTeamTagReport$((Get-Date -format yyyy-MMM-dd-ddd_hh-mm-ss_tt).ToString()).csv"
        Write-Progress -Activity "Getting reports...⏳" -Status "Please wait..."
        Get-MgTeam -All -ErrorAction Stop | ForEach-Object { $Team = $_ 
            Write-Progress -Activity "Getting TeamTagReport.." -Status "Fetching data for $($Team.DisplayName)" 
            $TeamMemberCheck = $true
            Get-MgTeamTag -All -TeamId $Team.id -ErrorAction Stop | ForEach-Object { $TeamTag = $_
                
                if ($TeamMemberCheck) {
                    $TeamMembersHashSet = @{}
                    Get-MgTeamMember -All -TeamId $Team.id -ErrorAction Stop | ForEach-Object {
                        $TeamMembersHashSet[$_.AdditionalProperties["userId"]] = $_.DisplayName }
                    $TeamMemberCheck = $false
                }
                $Output = [PSCustomObject]@{
                    "Team Name"          = $Team.DisplayName
                    "Team Member Count"  = $TeamMembersHashSet.Count 
                    "Team Description"   = $Team.Description 
                    "Team Visibility"    = $Team.Visibility 
                    "Tag Name"           = $TeamTag.DisplayName  
                    "Tag Member Count"   = $TeamTag.MemberCount
                    "Tag Type"           = $TeamTag.TagType
                    "Tag Members Id"         = ""
                    "Tag Members Teamwork Tag Id"    = ""
                    "NonTag Members Count" = 0
                    "NonTag Members Id"      = ""
                    "Team Id"            = $Team.id 
                    "Tag Id"             = $TeamTag.Id
                }
                $TagMembersHashSet = @{}

                Get-MgTeamTagMember -All -TeamId $Team.Id  -TeamworkTagId $TeamTag.Id -ErrorAction Stop | ForEach-Object {
                    $Output."Tag Members Id" += "$($_.DisplayName) ($($_.UserId)) ,`n"
                    $Output."Tag Members Teamwork Tag Id" += "$($_.DisplayName) ($($_.Id)),`n"
                    $TagMembersHashSet[$_.UserId] = $_.DisplayName
                }  
                ForEach ($TeamMember in $TeamMembersHashSet.Keys) {

                    if ($TagMembersHashSet[$TeamMember] -eq $null) {
                        $Output."NonTag Members Count"++
                        $Output."NonTag Members Id"  += "$($TeamMembersHashSet[$TeamMember]) ($($TeamMember)),`n"
                    }

                }
                $Output | Export-Csv -Path $FilePath -NoTypeInformation -Append -Force  
                
            }
        }
        InvokeCSVFile -ExportCSV $FilePath 
    } 

    catch {
        Write-Host "Error occurred: $( $_.Exception.Message )" -ForegroundColor Red
        Disconnect-MgGraph | Out-Null
        Exit
    }

}
function New-TeamTagMember {
   try{
      $NewUser=New-MgTeamTagMember -TeamId $TeamId.Trim() -TeamworkTagId $Teamworktagid.Trim() -UserId $UserId.Trim() -ErrorAction Stop
      $NewUser=$NewUser.DisplayName
      WriteToLogFile "$NewUser is added to the given TagId:$Teamworktagid successfully"
    }
    catch{
     WriteToLogFile  "$UserId was not added to the given TagId: $Teamworktagid 
     `n Error occurred: $( $_.Exception.Message )"
    }
    
}
function Remove-TeamTagMember {
  try{
     Remove-MgTeamTagMember -TeamId $TeamId.Trim() -TeamworkTagId $Teamworktagid.Trim() -TeamworkTagMemberId $UserTeamworkTagMemberId -ErrorAction Stop
     WriteToLogFile "$UserTeamworkTagMemberId is removed from the given TagId:$Teamworktagid successfully"
    }
    catch{
       WriteToLogFile  "$UserTeamworkTagMemberId is not removed from the  given TagId:$Teamworktagid 
       `n Error occurred: $( $_.Exception.Message )"
    }
}



function Add-UsertoTeamTag {
    
    if ($TeamId.Trim() -eq "") {
        $TeamId = Get-UserInput -Prompt "Enter the Team ID:"
    }
    if ($TeamworkTagId.Trim() -eq "") {
        $TeamworkTagId = Get-UserInput -Prompt "Enter the  Tag ID:"
    }

    $TeamMembersHashSet = @{}
    try{
    Get-MgTeamMember -All -TeamId $TeamId.Trim() -ErrorAction Stop | ForEach-Object {
        $TeamMembersHashSet[$_.AdditionalProperties["userId"]] = $_.DisplayName }}
          catch{
         Write-Host "Error occurred : $( $_.Exception.Message )" -ForegroundColor Red
         Disconnect-MgGraph | Out-Null
         Exit 
    
    }


  if($UserId.Trim() -eq  "" -and $CSV.Trim() -eq ""){
    Write-Host "`nOptions for Adding User to the Tag`n" -ForegroundColor Cyan
    Write-Host "    1. Single User using CommadLine"
    Write-Host "    2. Mulitple Users using CSV file`n"

    $Action = Get-UserInput -Prompt "Enter your choice:"}
    if($UserId.Trim() -ne  ""){
     $Action=1
    }
    if($CSV.Trim() -ne "")
     { $Action=2}
    


    switch ($Action) {
        '1' {
            if ($UserId.Trim() -eq "") {
                $UserId = Get-UserInput -Prompt "Enter the Userid:" 
            }
            if ($TeamMembersHashSet[$UserId.Trim()] -ne $null) {
                New-TeamTagMember
            }
            else {
                 WriteToLogFile "Given Userid $UserId not present in given TeamId $TeamId" #log
            }
        }
        '2' {
            if ($CSV.Trim() -eq "") {
                $CSV = Get-UserInput -Prompt "Enter the CSV Path:"  
            } 
            if ((Test-Path -Path $CSV) -eq "True") {
                
                Import-Csv -Path $CSV | ForEach-Object {
                
                    $UserId = $_.UserId  
                    if ($TeamMembersHashSet[$UserId.Trim()] -ne $null) {
                        New-TeamTagMember
                    } 
                else {
                    WriteToLogFile "Given Userid $UserId not present in given TeamId $TeamId" #log
                }}
            }
            else {
              
                Write-Host "Incorrect Csv File Path : $CSV" -ForegroundColor Red
              
            } 
               
        }
        default {
            Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
            Disconnect-MgGraph | Out-Null
            Exit
        }
    }
}
function Remove-UserFromTeamTag {
    if ($TeamId.Trim() -eq "") {
        $TeamId = Get-UserInput -Prompt "Enter the Team ID:"
    }
    if ($TeamworkTagId.Trim() -eq "") {
        $TeamworkTagId = Get-UserInput -Prompt "Enter the  Tag ID:"
    }
    if($UserTeamworkTagMemberId.Trim() -eq "" -and $CSV.Trim() -eq "" ){
    Write-Host "`nOptions for Removing User From the Tag`n" -ForegroundColor Cyan
    Write-Host "   1. Single User using CommadLine"
    Write-Host "   2. Mulitple Users using CSV file`n"

    $Action = Get-UserInput -Prompt "Enter your choice:"
    }
    if($UserTeamworkTagMemberId.Trim() -ne "")
    {
    $Action=1
    }
    if($CSV.Trim() -ne "")
    {
    $Action=2
    }
    switch ($Action) {
        '1' {
            if ($UserTeamworkTagMemberId.Trim() -eq "") {
                $UserTeamworkTagMemberId = Get-UserInput -Prompt "Enter the User TeamworkTagMemberId:"
            }
            Remove-TeamTagMember    
        }
        '2' {
            if ($CSV.Trim() -eq "") {
                $CSV = Get-UserInput -Prompt "Enter the CSV Path:"  
            } 
            if ((Test-Path -Path $CSV) -eq "True") {
                Import-Csv -Path $CSV | ForEach-Object {
                    $UserTeamworkTagMemberId = $_.TeamworkTagMemberId
                    Remove-TeamTagMember  
                }
            }
            else {
              
                Write-Host "Incorrect Csv File Path : $CSV" -ForegroundColor Red
              
            } 
     
        }
        default {
            Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
            Exit
        }
    }
}

#Connect to Microsoft.Graph.Teams

$Module = Get-Module -Name Microsoft.Graph.Teams -ListAvailable    
if ($Module.count -eq 0) {
    Write-Host Microsoft.Graph.Teams is not available in your System  -ForegroundColor Red
    $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No
    if ($Confirm -match "[yY]") {
        try {
            Install-Module Microsoft.Graph.Teams -Force -AllowClobber -Scope CurrentUser 
            Write-Host Microsoft.Graph.Teams installed successfully...  -ForegroundColor Green
        }
        catch {
            Write-Host "Error occurred : $( $_.Exception.Message )" -ForegroundColor Red
            Exit
        }
    }

    else {
        Write-Host Microsoft.Graph.Teams is required .Please Install-Module Microsoft.Graph.Teams to continue..
        Exit
    }
}

if ($TenantId -eq "") {
    $TenantId = Get-UserInput -Prompt "Enter the Tenant ID:"
}
if ($ClientID -eq "") {
    $ClientID = Get-UserInput -Prompt "Enter the Client ID:"
}
if ($CertificateThumbPrint -eq "") {
    $CertificateThumbPrint = Get-UserInput -Prompt "Enter the CertificateThumbPrint:"
}

if (($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbPrint -ne "")) {

    try {
        $Connect = Connect-MgGraph -TenantId $TenantID.Trim() -ClientID $ClientID.Trim() -CertificateThumbprint $CertificateThumbPrint.Trim() -ErrorAction Stop

        if ($Connect -ne $null) {
            Write-Host Microsoft.Graph.Teams Connected successfully... -ForegroundColor Green

            function Show-Menu {
                Write-Host "`n========================================="
                Write-Host "    📋 Team Tag Management " -ForegroundColor Cyan
                Write-Host "========================================="
                Write-Host "   1.Get Team Tags Reports"
                Write-Host "   2.Add Users to a TeamTag"
                Write-Host "   3.Remove Users From a TeamTag"
                Write-Host "=========================================`n"
            }
            function Execute-Choice {
                param([string]$Choice)
                switch ($Choice) {
                    '1' {
                        Get-TeamTagReports
                    }
                    '2' {
                         Add-UsertoTeamTag
                         InvokeLogFile -ExportCSV $logfile
                    }
                    '3' {
                        Remove-UserFromTeamTag
                        InvokeLogFile -ExportCSV $logfile
                    }
                    default {
                        Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
                        Disconnect-MgGraph | Out-Null
                        Exit
                    }

                }
            }

            if ($GetTeamTagReports) {
                Get-TeamTagReports
            }
            if ($AddUsertoTag) {
                 Add-UsertoTeamTag
                 InvokeLogFile -ExportCSV $logfile
                              } 
            
            if ($RemoveUserFromTag) {
                 Remove-UserFromTeamTag
                 InvokeLogFile -ExportCSV $logfile
            }
            if (-not ($GetTeamTagReports -or $AddUsertoTag -or $RemoveUserFromTag)) {
                Show-Menu
                $Choice = Get-UserInput -Prompt "Choose an action:"
                Execute-Choice $Choice
            }

        }

    }
    catch {
        Write-Host "Error occurred while connecting to Microsoft Graph: $( $_.Exception.Message )" -ForegroundColor Red
        Exit
    }
}
Disconnect-MgGraph | Out-Null
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green 
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n
