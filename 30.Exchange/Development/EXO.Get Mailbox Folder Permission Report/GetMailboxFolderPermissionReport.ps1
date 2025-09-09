 <#
=============================================================================================
Name: Get Mailbox Folder Permission Report Using PowerShell
Version: 1.0
Website: o365reports.com

~~~~~~~~~~~~~~~~~
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script automatically installs the Exchange PowerShell module (if not installed already) upon your confirmation. 
2. The script can generate 7+ folder permission reports. 
3. Retrieves all mailbox folders and their permissions for all mailboxes. 
4. Shows permission for a specific folder in all mailboxes. 
5. Get a list of mailbox folders a user has access to. 
6. Retrieves all mailbox folders delegated with specific access rights. 
7. Provides option to exclude default and anonymous access. 
8. Allows to get folder permissions for all user mailboxes. 
9. Allows to get folder permissions for all shared mailboxes. 
10. Exports report results to CSV. 
11. The script is scheduler friendly. 
12. It can be executed with certificate-based authentication (CBA) too.

For detailed script execution: https://o365reports.com/2024/06/05/get-mailbox-folder-permission-report-using-powershell/

============================================================================================
#>   
Param (
   [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [string]$Organization,
        [string]$CertificateThumbprint,
        [string]$UserName,
        [string]$Password,
        [string]$MailboxUPN ,
        [string]$MailboxCSV  ,
        [string]$SpecificFolder ,
        [string]$FoldersUserCanAccess,
        [ValidateSet("None","Reviewer","PublishingEditor","PublishingAuthor","Owner","NonEditingAuthor","Editor","Contributor","Author")]
        [array]$AccessRights,
        [switch]$ExcludeDefaultAndAnonymousUsers,
        [switch]$UserMailboxOnly ,
        [switch]$SharedMailboxOnly
       
      
       )
   
    #Check for EXO module inatallation
    $Module = Get-Module ExchangeOnlineManagement -ListAvailable
    if($Module.count -eq 0) 
    { 
     Write-Host Exchange Online PowerShell  module is not available  -ForegroundColor yellow  
     $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
    if($Confirm -match "[yY]") 
    { 
     Write-host "Installing Exchange Online PowerShell module"
     Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
    } 
    else 
    { 
     Write-Host EXO module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
     Exit
    }
   } 
    Write-Host Connecting to Exchange Online...
      if(($UserName -ne "") -and ($Password -ne ""))
    {
    $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
    $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
    Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
    }
    elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
    {
    Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
    }
    else
    {
    Connect-ExchangeOnline -ShowBanner:$false
    }
    
#Output file declaration 
$Location = (Get-Location) 
$OutputCSV = "$($Location)\MailboxFolderPermissionReport_$((Get-Date -format 'yyyy-MMM-dd-ddd hh-mm-ss').ToString()).csv"

#Function to export csv in mailbox folder permissions
Function GetPermission {
    param($MailboxUPN, $FolderPermissions)

        foreach ($Permission in $FolderPermissions) {
            $MailboxFolderPermissionsData = [PSCustomObject]@{
                "Display Name" =$DisplayName
                "UPN" = $MailboxUPN
                "Mailbox Type" = $MailboxType
                "Folder Name" = $Permission.FolderName
                "Folder Identity" = $Permission.Identity
                "Shared To" = $Permission.User
                "Access Rights" =$Permission.AccessRights.Trim('{', '}')
                 }  
                 $MailboxFolderPermissionsData | Export-Csv -Path "$OutputCSV" -Append -NoTypeInformation -Force
        }
           
 } 
 
#Function to get mailbox folder names
Function FolderStatistics{  
    param($MailboxUPN)
          
        Get-EXOMailboxFolderStatistics -Identity $MailboxUPN | Foreach {  
        $FolderIdentity =$_.Identity
        $FolderName =$FolderIdentity.Substring($FolderIdentity.IndexOf('\')+1)
            if($FolderName -like "Top of Information Store"){
            $FolderName =""
            } 
            elseif($FolderName -in "Recoverable Items","Audits","Calendar Logging","Deletions","Purges","Versions","SubstrateHolds","DiscoveryHolds"){
                   return 
            }
            if($ExcludeDefaultAndAnonymousUsers.IsPresent){
               ProcessExcludeDefaultAndAnonymousUsers -MailboxUPN $MailboxUPN -Foldername $FolderName
            }
            elseif($FoldersUserCanAccess -ne ""){
               ProcessFoldersUserCanAccess -MailboxUPN $MailboxUPN -Foldername $FolderName -User $FoldersUserCanAccess
            }
            elseif($AccessRights.Count -gt 0) {
               ProcessToFilterFolderPermissionsByAccessRights -MailboxUPN $MailboxUPN -AccessRights $AccessRights -Foldername $FolderName
            }
            else{
               ProcessAllMailboxFolderPermission -MailboxUPN $MailboxUPN -Foldername $FolderName
            }
         }
} 

#Function to permission for all mailbox folders
Function ProcessAllMailboxFolderPermission {
   param($MailboxUPN,$Foldername)
     $FolderPermissions = Get-EXOMailboxFolderPermission -Identity "${MailboxUPN}:\${FolderName}" 
     GetPermission -MailboxUPN $MailboxUPN  -FolderPermissions $FolderPermissions   
}  

#Function to mailbox permissions for particular folders
Function ProcessSpecificMailboxFolderPermission {
 param($MailboxUPN, $FolderName)
       $FolderPermissions = Get-EXOMailboxFolderPermission -Identity "${MailboxUPN}:\${FolderName}"
        if($FolderPermissions){
           GetPermission -MailboxUPN $MailboxUPN -FolderPermissions $FolderPermissions
        }
        else{ 
        Write-Host "Failed to Specific folder statistics for mailbox: $MailboxUPN IN $SpecificFolder" -ForegroundColor Yellow
        }  
} 
 
# Function to mailbox folders user permission 
Function ProcessExcludeDefaultAndAnonymousUsers{
  param($MailboxUPN,$FolderName )
     $FolderPermissions = Get-EXOMailboxFolderPermission -Identity "${MailboxUPN}:\${FolderName}"| Where-Object { $_.User -notin @("Default", "Anonymous") }
       GetPermission -MailboxUPN $MailboxUPN  -FolderPermissions $FolderPermissions

 }
#Function to Identify the particular user in the mailbox folders permission          
Function ProcessFoldersUserCanAccess{
   param($MailboxUPN,$FolderName ,$User)        
          $FolderPermissions = Get-EXOMailboxFolderPermission -Identity "${MailboxUPN}:\${FolderName}" -User $User -ErrorAction SilentlyContinue
          GetPermission -MailboxUPN $MailboxUPN  -FolderPermissions $FolderPermissions
}

#Function to retrieve permissions for all folders with a specific access right
Function ProcessToFilterFolderPermissionsByAccessRights {
   param(
        $MailboxUPN,
        $AccessRights,
        $FolderName
        )
       $FolderPermissions = Get-MailboxFolderPermission -Identity "${MailboxUPN}:\${FolderName}"| Where-Object { $_.AccessRights -in $AccessRights}
       if($FolderPermissions){ 
       
       GetPermission -mailboxUPN $MailboxUPN -FolderPermissions $FolderPermissions      
        }
}
#Function to get Mailbox
Function Getmailbox{
   Param($MailboxUPN)
    
     $MailBoxInfo = Get-EXOMailbox -UserPrincipalName $MailboxUPN
       $DisplayName =$MailboxInfo.DisplayName
       $MailboxType = $MailboxInfo.RecipientTypeDetails
        Process-Mailbox -MailboxUPN $MailboxUPN
}

Function Process-Mailbox{
     Param ($MailboxUPN)
    Write-Progress -Activity "Processed Mailbox Count  : $ProgressIndex" -Status "Currently Processing : $MailboxUPN"    
    if($SpecificFolder -ne ""){
         ProcessSpecificMailboxFolderPermission -MailboxUPN $MailboxUPN -FolderName $SpecificFolder
     }
     else{
         FolderStatistics -MailboxUPN $MailboxUPN
     }
}

#Single mailboxUPN
if ($MailboxUPN) {
               $ProgressIndex =1
               Getmailbox -MailboxUPN $MailboxUPN     
} 
#CSV file input
elseif($MailBoxCSV) {
           
        $Mailboxes = Import-Csv -Path $MailBoxCSV
        $ProgressIndex = 0
        foreach ($Mailbox in $Mailboxes) {
        $ProgressIndex++
        Getmailbox -MailboxUPN $Mailbox.Mailboxes
        }
 }

else{ 
    $ProgressIndex =0
    if($SharedMailboxOnly.IsPresent -or $UserMailboxOnly.IsPresent){
  
         if ($SharedMailboxOnly.IsPresent) {
           $RecipientType ="SharedMailbox"
         }
         else {
           $RecipientType = "UserMailbox"
         }
           Get-EXOMailbox -RecipientTypeDetails $RecipientType  -ResultSize Unlimited| Foreach {
           $ProgressIndex++
           $DisplayName =$_.DisplayName
           $MailboxUPN = $_.UserPrincipalName
           $MailboxType = $_.RecipientTypeDetails
           Process-Mailbox -MailboxUPN $MailboxUPN
           }
    }     
    else{ 
           Get-EXOMailbox -ResultSize Unlimited| Foreach {
           $ProgressIndex++
           $DisplayName =$_.DisplayName
           $MailboxUPN = $_.UserPrincipalName
           $MailboxType = $_.RecipientTypeDetails
           Process-Mailbox -MailboxUPN $MailboxUPN
           }
     }     
}
if (Test-Path -Path $OutputCSV) {     
     $ItemCounts = (Import-Csv -Path "$OutputCSV").Count
    
     if($ItemCounts -ne 0){
           
            Write-Host "`nThe output file contains $($ProgressIndex) mailboxes and $($ItemCounts) records." -ForegroundColor Cyan
     
            Write-Host "`nThe output file is available at:" -ForegroundColor Yellow ;Write-Host $OutputCSV -ForegroundColor Cyan
            Write-Host "`n~~ Script prepared by AdminDroid Community ~~`n" -ForegroundColor Green

            Write-Host "~~ Check out " -NoNewline -ForegroundColor Green
            Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline
            Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green
            Write-Host "`n`n"
            
            $Prompt = New-Object -ComObject wscript.shell
            $UserInput = $Prompt.popup("Do you want to open the output file?", 0, "Open Output File", 4)

            if ($UserInput -eq 6) {
                Invoke-Item "$OutputCSV"
            }
      }
}       
else {
      Write-Host "No records found" -ForegroundColor Yellow
}
        
#Disconnect Exchange Online session
 Disconnect-ExchangeOnline -Confirm:$false        
     