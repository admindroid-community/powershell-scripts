<#
=============================================================================================
Name:           Copy Members from One Distribution List to Another 
Version:        1.0
Website:        o365reports.com


Script Highlights:  
~~~~~~~~~~~~~~~~~
1. Fetches members and owners from the source DL and adds them to the target DL.   
2. Helps to copy DL members from one to another. (if not present already). 
3. Can copy managers from one DL to another. (if not present already). 
4. The script automatically verifies and installs the Exchange Online PowerShell module (if not installed) upon your confirmation.  
5. Exports an output log file in TXT format that shows the result of the operation.   
6. The script can be executed with MFA enabled account too.  
7. The script supports certificate-based authentication (CBA) too. 
8. The script is schedular-friendly.   

For detailed Script execution: https://o365reports.com/2025/04/08/copy-members-from-one-distribution-list-to-another/


============================================================================================
#>
Param
(
    [Parameter(Mandatory = $False)]
	[string]$SourceGroup=$Null,
	[string]$TargetGroup=$Null,
	[Switch]$CopyOwnersOnly,
	[Switch]$CopyMembersOnly,
    [string]$UserName=$Null,
    [string]$Password=$Null,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbPrint
)

#Connection module
Function Connect_Module {
    #Check for Exchange Online module installation
    $ExchangeModule = Get-Module ExchangeOnlineManagement -ListAvailable
    if($ExchangeModule.count -eq 0) {
        Write-Host ExchangeOnline module is not available -ForegroundColor Yellow
        $confirm = Read-Host Do you want to Install ExchangeOnline module? [Y] Yes  [N] No
        if($confirm -match "[Yy]") {
            Write-Host "Installing ExchangeOnline module ..."
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
            Import-Module ExchangeOnlineManagement
        }    
        else {
            Write-Host "ExchangeOnline Module is required. To Install ExchangeOnline module use 'Install-Module ExchangeOnlineManagement' cmdlet."
            Exit
        }
    }

    #Connect Exchange Online
    Write-Host "`nConnecting Exchange Online module ..."
    if(($UserName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
        Connect-ExchangeOnline -Credential $Credential 
    }
    elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "") {
        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
}

#Output log function
Function Log-Data {
    param ([string]$UserName,$UserType,$Status,$Message = "")
    $Result = if ($Status -eq "Failed") {
        "- Failed to copy $UserName as a $UserType to the target group. Error: $Message"
    } elseif(($Status -eq "Already a Member") -or ($Status -eq"Already a Owner")){
        "- Skipped adding $UserName as $UserType. The user is already a $UserType in the target group."
    } else {
        "- $UserName has been successfully copied as $UserType to the target group."
    }
    Add-Content -Path $OutputLog -Value $Result
}

$Script:CopiedMembersCount = 0
$Script:CopiedOwnersCount = 0
$Script:ExistingMembersCount = 0
$Script:ExistingOwnersCount = 0

#Funtion to copy members only to another distribution group
Function Copy_Members
{
	$ProcessedUserCount=0

    $SourceGroupMemberNames = Get-DistributionGroupMember -Identity $SourceGroup
    $TargetGroupMemberNames = Get-DistributionGroupMember -Identity $TargetGroup

    ForEach($SourceGroupMemberName in $SourceGroupMemberNames)
    {
        $SourceMemberName = $SourceGroupMemberName.DisplayName
        Write-Progress -Activity "Copying Members From $SourceGroup to $TargetGroup " -Status "Processed member Count : $ProcessedUserCount" -CurrentOperation " Currently Processing Member $($SourceMemberName)"
        Try
        {
            if($SourceGroupMemberName.Guid -in $TargetGroupMemberNames.Guid){
                Log-Data -UserName $SourceGroupMemberName.DisplayName -UserType "member" -Status "Already a Member"
                $Script:ExistingMembersCount++
            }
            else{
                Add-DistributiongroupMember -Identity $TargetGroup -Member $($SourceGroupMemberName.DistinguishedName) -ErrorAction Stop
                Log-Data -UserName $SourceGroupMemberName.DisplayName -UserType "member" -Status "Successfully Copied"
                $Script:CopiedMembersCount++
            }
        }
        Catch
        {
            Log-Data -UserName $Member.DisplayName -UserType "member" -Status "Failed" -Message $_.Exception.Message
        }
		$ProcessedUserCount++
    } 
}

#Funtion to copy owners only to another distribution group
Function Copy_Owners
{	   
	$ProcessedUserCount=0

    $SourceGroupOwnerNames = Get-DistributionGroup -Identity $SourceGroup | Select -ExpandProperty ManagedBy
    $TargetGroupOwnerNames = Get-DistributionGroup -Identity $TargetGroup | Select -ExpandProperty ManagedBy

    ForEach($SourceGroupOwnerName in $SourceGroupOwnerNames)
    {
        $Owner = Get-Recipient -Identity $SourceGroupOwnerName
		Write-Progress -Activity "Copying owner From $SourceGroup To $TargetGroup " -Status "Processed owner Count: $ProcessedUserCount" -CurrentOperation "Currently Processing Owner $($Owner.Displayname)"
        try{
            if($SourceGroupOwnerName -in $TargetGroupOwnerNames)
            {
	            Log-Data -UserName $Owner.DisplayName -UserType "owner" -Status "Already a Owner"
                $Script:ExistingOwnersCount++
		    }
            else
            {
                Set-DistributionGroup -Identity $TargetGroup -ManagedBy  @{Add="$($Owner.PrimarySmtpAddress)"} -ErrorAction Stop
	            Log-Data -UserName $Owner.DisplayName -UserType "owner" -Status "Successfully Copied"
                $Script:CopiedOwnersCount++
            }
        }
        catch{
            Log-Data -UserName $Owner.DisplayName -UserType "owner" -Status "Failed" -Message $_.Exception.Message
        }
		$ProcessedUserCount++
    }
}

#Function to return output log file
Function OpenOutputLog
{  		
    Write-Host "`nSummary of Distribution List Membership Changes" -ForegroundColor DarkYellow
    Write-Host "`n$($ExistingMembersCount) member(s) and $($ExistingOwnersCount) owner(s) from source DL is already present in target DL."
    #Write-Host "`n$($ExistingMembersCount) members and $($ExistingOwnersCount) owners already present in both distribution lists."
    Write-Host "`nThe script successfully copied $($CopiedMembersCount) members and $($CopiedOwnersCount) owners to the target DL."

    #Open Output file after execution 
    If((Test-Path -Path $OutputLog) -eq "True") 
    {			
        Write-Host `nThe log file available in :`n -NoNewline -ForegroundColor Yellow
        Write-Host $OutputLog
        Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
        Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
        $Prompt = New-Object -ComObject wscript.shell    
        $UserInput = $Prompt.popup("Do you want to open log file?",` 0,"Open log File",4)    
        If($UserInput -eq 6)    
        {    
            Invoke-Item "$OutputLog"    
        }  
    }    	
}

#Execution starts here
Connect_Module
If($SourceGroup -eq "" -or ($TargetGroup -eq ""))
{
	If($SourceGroup -eq "")
	{
		$SourceGroup = Read-Host "`nEnter the source DL mail Id"
	}
	If($TargetGroup -eq "")
	{
		$TargetGroup = Read-Host "`nEnter the target DL mail Id"
	}
}

$Location = Get-Location
$OutputLog="$Location\Copy_DL_Membership_Log_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).txt"
If($CopyOwnersOnly.IsPresent)
{
	Copy_Owners
}
ElseIf($CopyMemberOnly.IsPresent)
{
	Copy_Members
}
Else
{
	Copy_Owners
	Copy_Members
}

#Open output log file
OpenOutputLog

#Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false | Out-Null