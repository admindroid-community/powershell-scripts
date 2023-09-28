<#
=============================================================================================
Name:           Export Distrbution Groups a user is member of
Description:    This script exports all users and their distribution group membership
Website:        o365reports.com

Script Highlights:
~~~~~~~~~~~~~~~~~
1. The script uses modern authentication to connect to Exchange Online.
2. The script can be executed with MFA enabled account
3. Automatically installs the EXO V2 module (if not installed already) upon your confirmation.
4. Credentials are passed as parameters, so worry not!
5. Allows generating user membership reports based on your requirement.
   a.DL membership for all users.
   b. DL membership for a list of users (import CSV).
   c. DL membership for a single user.

For detailed script execution: https://o365reports.com/2022/04/19/list-all-the-distribution-groups-a-user-is-member-of-using-powershell/
============================================================================================
#>
Param
(
    [string]$UserName=$Null,
    [string]$Password=$Null,
	[string]$UserPrincipalName=$Null,
	[string]$InputCsvFilePath=$Null
)
Function Connect_Exo
{		
    #Check for EXO v2 module inatallation
    $Module = Get-Module ExchangeOnlineManagement -ListAvailable
    If($Module.count -eq 0) 
    {			
        Write-Host "Exchange Online PowerShell V2 module is not available"  -ForegroundColor yellow  
        $Confirm= Read-Host "Are you sure you want to install module? [Y] Yes [N] No" 
        If($Confirm -match "[yY]")			   
        {						
            Write-host "Installing Exchange Online PowerShell module"
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
        } 
        Else 
        { 
            Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
            Exit
        }
    }   
    #Importing Module by default will avoid the cmdlet unrecognized error
    Import-Module ExchangeOnlineManagement
    Write-Host Connecting to Exchange Online...
    #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
    If(($UserName -ne "") -and ($Password -ne ""))
    {			
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
        Connect-ExchangeOnline  -Credential $Credential
    }
    Else
    {
        Connect-ExchangeOnline
    }
} 

Function Import_Csv
{
    #Importing UserPrincipalName From The Csv
    Try
    {
        $UserDetails=@()
        Write-Host "Importing UserPrincipalNames from Csv..."
        $UPNs=Import-Csv $InputCsvFilePath 
        foreach ($UPN in $UPNs)
        {
        $UserPrincipalName=$UPN.User_Principal_Name
            Try
	        {   
		         Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop |foreach{
                     List_DLs_That_User_Is_A_Member
                   
	        }
}
	        Catch
            {
		        Write-Host "$UserPrincipalName is not a valid user"
	        }
        }

    }
    catch
    {
        Write-Host "$InputCsvFilePath is not a valid file path"
    }     
}
Function List_DLs_That_User_Is_A_Member
{
    #Finding Distribution List that  User is a Member
	$Result= @()
    $DistinguishedName=$_.DistinguishedName
    $Filter = "Members -Like ""$DistinguishedName"""
    $UserPrincipalName=$_.UserPrincipalName
    $UserDisplayName=$_.DisplayName
    Write-Progress -Activity "Find Distribution Lists that user is a member" -Status "Processed User Count: $Global:ProcessedUserCount" -CurrentOperation "Currently Processing in  $UserPrincipalName"
    $DLs=Get-DistributionGroup -ResultSize Unlimited -Filter $Filter
    $GroupCount=$DLs | Measure-Object | select count
    If($GroupCount.count -ne 0)
    {    
	    $DLsCount=$GroupCount.count
		$DLsName=$DLs.Name
	    $DLsEmailAddress=$DLs.PrimarySmtpAddress
    }
    Else
    {
	    $DLsName="-"
	    $DlsEmailAddress="-"
		$DLsCount='0'
    }
    $Result=New-Object PsObject -Property @{'User Principal Name'=$UserPrincipalName;'User Display Name'=$UserDisplayName;'No of DLs that user is a member'=$DLsCount;'DLs Name'=$DLsName -join ',';'DLs Email Adddress'=$DLsEmailAddress -join ',';} 
    $Result|Select-Object 'User Principal Name','User Display Name','No Of DLs That User Is A Member','DLs Name','DLs Email Adddress'| Export-Csv  $OutputCsv -NoTypeInformatio -Append 
    $Global:ProcessedUserCount++		
  
}
Function OpenOutputCsv
{  		
    #Open Output File After Execution 
    If((Test-Path $OutputCsv) -eq "True") 
    {			
        Write-Host `n"The output file contains:" -NoNewline -ForegroundColor Yellow; Write-Host $ProcessedUserCount users `n
        Write-Host " The Output file available in:" -NoNewline -ForegroundColor Yellow; $OutputCsv
        Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
        Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline;                                            Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
        $Prompt = New-Object -ComObject wscript.shell    
        $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"open output file",4)    
        If($UserInput -eq 6)    
        {    
            Invoke-Item "$OutputCsv"    
        }  
    } 	
}
Connect_Exo
$Global:ProcessedUserCount=1
$OutputCsv=".\ListDLs_UsersIsMemberOf_$((Get-Date -format MMM-dd` hh-mm` tt).ToString()).csv"
If($UserPrincipalName -ne "")
{  
	Try
	{
        write-Host "Checking $UserPrincipalName is a valid user or not"
		Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop|ForEach{
            List_DLs_That_User_Is_A_Member
        }
	}
	Catch
    {
		Write-Host "$UserPrincipalName is not a valid user"
	}
}		
Elseif($InputCsvFilePath -ne "")
{	
    Import_Csv
}
Else
{ 
    Get-Mailbox -ResultSize unlimited -RecipientTypeDetails UserMailbox | ForEach{
	    List_DLs_That_User_Is_A_Member
    }
}
OpenOutputCsv
#Removing connected session
Get-PSSession |Remove-PSSession