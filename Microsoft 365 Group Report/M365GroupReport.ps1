<#
=============================================================================================
Name:           Microsoft 365 Group Report
Description:    This script exports Microsoft 365 groups and their membership to CSV using Microsoft Graph PowerShell
Version:        3.0
Website:        o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1.The script uses Microsoft Graph PowerShell.
2.The script can be executed with certificate-based authentication (CBA) too.
3.Exports the report result to CSV. 
4.You can get members count based on Member Type such as User, Group, Contact, etc. 
5.The script is scheduler friendly.
6.Above all, the script exports output to nicely formatted 2 CSV files. One with group information and another with detailed group membership information. 

For detailed Script execution: https://o365reports.com/2021/02/11/export-microsoft-365-group-report-to-csv-using-powershell
============================================================================================
#>

Param 
( 
    [Parameter(Mandatory = $false)] 
    [string]$GroupIDsFile,
    [switch]$DistributionList, 
    [switch]$Security, 
    [switch]$MailEnabledSecurity, 
    [Switch]$IsEmpty, 
    [Int]$MinGroupMembersCount,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
) 

Function Get_members
{
    $DisplayName=$_.DisplayName
    Write-Progress -Activity "`n     Processed Group count: $Count "`n"  Getting members of: $DisplayName"
    $EmailAddress=$_.Mail
    if($_.GroupTypes -eq "Unified")
    {
        $GroupType="Microsoft 365"
    }
    elseif($_.Mail -ne $null)
    {
        if($_.SecurityEnabled -eq $false)
        {
            $GroupType="DistributionList"
        }
        else
        {
            $GroupType="MailEnabledSecurity"
        }
    }
    else
    {
        $GroupType="Security"
    }
    $GroupId=$_.Id
    $Recipient=""
    $RecipientHash=@{}
    for($KeyIndex = 0; $KeyIndex -lt $RecipientTypeArray.Length; $KeyIndex += 2)
    {
        $key=$RecipientTypeArray[$KeyIndex]
        $Value=$RecipientTypeArray[$KeyIndex+1]
        $RecipientHash.Add($key,$Value)
    }
    $Members=Get-MgGroupMember -All -GroupId $GroupId
    $MembersCount=$Members.Count
    $Members=$Members.AdditionalProperties
    #Filter for security group
    if(($Security.IsPresent) -and ($GroupType -ne "Security"))
    {
        Return
    }

    #Filter for Distribution list
    if(($DistributionList.IsPresent) -and ($GroupType -ne "DistributionList"))
    {
        Return
    }

    #Filter for mail enabled security group
    if(($MailEnabledSecurity.IsPresent) -and ($GroupType -ne "MailEnabledSecurity"))
    {
        Return
    }

    #GroupSize Filter
    if(([int]$MinGroupMembersCount -ne "") -and ($MembersCount -lt [int]$MinGroupMembersCount))
    {
        Return
    }
    #Check for Empty Group
    elseif($MembersCount -eq 0)
    {
        $MemberName="No Members"
        $MemberEmail="-"
        $RecipientTypeDetail="-"
        Print_Output
    }
    #Loop through each member in a group
    else
    {
        foreach($Member in $Members){
            if($IsEmpty.IsPresent)
            {
                return
            }
            $MemberName=$Member.displayName
            if($Member.'@odata.type' -eq '#microsoft.graph.user')
            {
                $MemberType="User"
            }
            elseif($Member.'@odata.type' -eq '#microsoft.graph.group')
            {
                $MemberType="Group"
            }
            elseif($Member.'@odata.type' -eq '#microsoft.graph.orgContact')
            {
                $MemberType="Contact"
            }
            $MemberEmail=$Member.mail
            if($MemberEmail -eq "")
            {
                $MemberEmail="-"
            }
            #Get Counts by RecipientTypeDetail
            foreach($key in [object[]]$Recipienthash.Keys){
                if(($MemberType -eq $key) -eq "true")
                {
                    [int]$RecipientHash[$key]+=1
                }
            }
            Print_Output
        }
    }
 
    #Order RecipientTypeDetail based on count
    $Hash=@{}
    $Hash=$RecipientHash.GetEnumerator() | Sort-Object -Property value -Descending |foreach{
        if([int]$($_.Value) -gt 0 )
        {
            if($Recipient -ne "")
            {
                $Recipient+=";"
            } 
            $Recipient+=@("$($_.Key) - $($_.Value)")    
        }
        if($Recipient -eq "")
        {
            $Recipient="-"
        }
    }
    #Print Summary report
    $Result=@{'DisplayName'=$DisplayName;'EmailAddress'=$EmailAddress;'GroupType'=$GroupType;'GroupMembersCount'=$MembersCount;'MembersCountByType'=$Recipient}
    $Results= New-Object PSObject -Property $Result 
    $Results | Select-Object DisplayName,EmailAddress,GroupType,GroupMembersCount,MembersCountByType | Export-Csv -Path $ExportSummaryCSV -Notype -Append
}

#Print Detailed Output
Function Print_Output
{
    $Result=@{'GroupName'=$DisplayName;'GroupEmailAddress'=$EmailAddress;'Member'=$MemberName;'MemberEmail'=$MemberEmail;'MemberType'=$MemberType} 
    $Results= New-Object PSObject -Property $Result 
    $Results | Select-Object GroupName,GroupEmailAddress,Member,MemberEmail,MemberType | Export-Csv -Path $ExportCSV -Notype -Append
}
Function CloseConnection
{
    Disconnect-MgGraph | Out-Null
    Exit
}
Function main() 
{
    #Check for MSOnline module 
    $MsGraphModule =  Get-Module Microsoft.Graph -ListAvailable  
    if($MsGraphModule -eq $null)
    { 
        Write-host "Important: MicrosoftGraph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No  
        if($confirm -match "[yY]") 
        { 
            Write-host `n"Installing MicrosoftGraph module..."
            Install-Module Microsoft.Graph -Repository PsGallery -Force -AllowClobber -Scope CurrentUser
            Write-host `n"Required Module is installed in the machine Successfully" -ForegroundColor Magenta 
        } 
        else
        { 
            Write-host "Exiting. `nNote: MsGraph module must be available in your system to run the script. Please install required module." -ForegroundColor Red 
            Exit 
        } 
    } 
    Write-Host `n"Connecting to Microsoft Graph..."`n
    $Scopes = @("Directory.Read.All"
    )  
    #Storing credential in script for scheduling purpose/ Passing credential as parameter
    $Error.Clear()  
    if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
    {  
        try
        {
            Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint 
        }
        catch
        {
            Write-Host `n"Please provide Correct Details!" -ForegroundColor Red
            Exit
        }
    }  
    else  
    {
        Connect-MgGraph -Scopes $Scopes
    } 
    Write-Host `n"Microsoft Graph connected" -ForegroundColor Green
    #Set output file 
    $ExportCSV=".\M365Group-DetailedMembersReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" #Detailed report
    $ExportSummaryCSV=".\M365Group-SummaryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" #Summary report

    #Get a list of RecipientTypeDetail
    $RecipientTypeArray=Get-Content -Path .\RecipientTypeDetails.txt -ErrorAction Stop
    $Result=""  
    $Results=@()
    $Count=0
    Write-Progress -Activity "Collecting group info"
    #Check for input file
    if([string]$GroupIDsFile -ne "") 
    { 
        #We have an input file, read it into memory 
        $DG=@()
        $DG=Import-Csv -Header "DisplayName" $GroupIDsFile
        foreach($item in $DG){
            Get-MgGroup -GroupId $item.displayname | Foreach{
                $Count++
                Get_Members
            }
        }
    }
    else
    {
        #Get all Office 365 group
        Get-MgGroup -All -ErrorAction SilentlyContinue -ErrorVariable PermissionError| Foreach{
            $Count++
            Get_Members
        }
        if($PermissionError)
        {
            Write-Host "Please Add permissions!" -ForegroundColor Red
            CloseConnection
        }
    }

    #Open output file after execution 
    Write-Host `n"Script executed successfully"
    if((Test-Path -Path $ExportCSV) -eq "True")
    {
        Write-Host `n" Detailed report available in:" -NoNewline -ForegroundColor Yellow
		Write-Host $ExportCSV 
        Write-host `n" Summary report available in:" -NoNewline -ForegroundColor Yellow
		Write-Host $ExportSummaryCSV 
		Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
        Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
        $Prompt = New-Object -ComObject wscript.shell  
        $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)  
        If ($UserInput -eq 6)  
        {  
            Invoke-Item "$ExportCSV"  
            Invoke-Item "$ExportSummaryCSV"
            CloseConnection
        } 
    }
    Else
    {
        Write-Host `n"No group found" -ForegroundColor Red
        CloseConnection
    }
}
. main

