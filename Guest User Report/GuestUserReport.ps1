<#
=============================================================================================
Name:           Export Office 365 Guest user and their membership report using MS Graph PowerShell
Version:        3.0
Website:        o365reports.com
For detailed Script execution: https://o365reports.com/2020/11/12/export-office-365-guest-user-report-with-their-membership/
============================================================================================
#>
#Accept input parameter
Param
(
    [Parameter(Mandatory = $false)]
    [int]$StaleGuests,
    [int]$RecentlyCreatedGuests,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

$MsGraphModule =  Get-Module Microsoft.Graph -ListAvailable
if($MsGraphModule -eq $null)
{ 
    Write-host "Important: Microsoft graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
    $confirm = Read-Host Are you sure you want to install Microsoft graph module? [Y] Yes [N] No  
    if($confirm -match "[yY]") 
    { 
        Write-host "Installing Microsoft graph module..."
        Install-Module Microsoft.Graph -Scope CurrentUser
        Write-host "Microsoft graph module is installed in the machine successfully" -ForegroundColor Magenta 
    } 
    else
    { 
        Write-host "Exiting. `nNote: Microsoft graph module must be available in your system to run the script" -ForegroundColor Red
        Exit 
    } 
}
if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
{  
    Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction SilentlyContinue -ErrorVariable ConnectionError|Out-Null
    if($ConnectionError -ne $null)
    {    
        Write-Host $ConnectionError -Foregroundcolor Red
        Exit
    }
}
else
{
    Connect-MgGraph -Scopes "Directory.Read.All"  -ErrorAction SilentlyContinue -Errorvariable ConnectionError |Out-Null
    if($ConnectionError -ne $null)
    {
        Write-Host "$ConnectionError" -Foregroundcolor Red
        Exit
    }
}
Write-Host "Microsoft Graph Powershell module is connected successfully" -ForegroundColor Green
Select-MgProfile beta
$Result=""   
$GuestCount=0
$PrintedGuests=0

#Output file declaration 
$ExportCSV=".\GuestUserReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
Write-Host `nExporting report... 
#Getting guest users
Get-MgUser -All -Filter "UserType eq 'Guest'" -ExpandProperty MemberOf  | foreach {
    $DisplayName = $_.DisplayName
    $GuestCount++
    Write-Progress -Activity "`n     Processed mailbox count: $GuestCount "`n"  Currently Processing: $DisplayName"
    $AccountAge = (New-TimeSpan -Start $_.CreatedDateTime).Days

    #Check for stale guest users
    if(($StaleGuests -ne "") -and ([int]$AccountAge -lt $StaleGuests)) 
    { 
        return
    }

    #Check for recently created guest users
    if(($RecentlyCreatedGuests -ne "") -and ([int]$AccountAge -gt $RecentlyCreatedGuests)) 
    { 
        return
    }
    $Company = $_.CompanyName
    if($Company -eq $null)
    {
        $Company = "-"
    }
    $GroupMembership = @($_.MemberOf.AdditionalProperties.displayName) -join ','
    if($GroupMembership -eq $null)
    {
        $GroupMembership = '-'
    }
    #Export result to CSV file 
    $PrintedGuests++
    $Result = [PSCustomObject] @{'DisplayName'=$DisplayName;'UserPrincipalName'=$_.UserPrincipalName;'Company'=$Company;'EmailAddress'=$_.Mail;'CreationTime'=$_.CreatedDateTime ;'AccountAge(days)'=$AccountAge;'CreationType'=$_.CreationType;'InvitationAccepted'=$_.ExternalUserState;'GroupMembership'=$GroupMembership} 
    $Result | Export-Csv -Path $ExportCSV -Notype -Append
}

#Open output file after execution 
Write-Host `nScript executed successfully
if((Test-Path -Path $ExportCSV) -eq "True")
{
    Write-Host "Detailed report available in: $ExportCSV" -ForegroundColor Green
    Write-Host `nThe Output file contains $PrintedGuests guest users.
    $Prompt = New-Object -ComObject wscript.shell  
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)  
    if ($UserInput -eq 6)  
    {  
        Invoke-Item "$ExportCSV"  
    } 
}
else
{
    Write-Host "No guest user found" -ForegroundColor Red
}
Disconnect-MgGraph|Out-Null