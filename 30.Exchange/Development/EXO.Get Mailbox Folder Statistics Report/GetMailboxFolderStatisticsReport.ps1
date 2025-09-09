<#
=============================================================================================
Name: Get Exchange Online Mailbox Folder Statistics Using PowerShell
Version:1.0
Website: o365reports.com

~~~~~~~~~~~~~~~~~
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script verifies and installs Exchange PowerShell module (if not installed already) upon your confirmation.
2. Retrieve folder statistics for all mailbox folders.
3. Retrieve statistics for specific mailbox folders.
4. Provides folder statistics for a single user and bulk users.
5. Allows to use filter to get folder statistics for all user mailboxes.
6. Allows to use filter to get folder statistics for all shared mailboxes.
7. The script can be executed with an MFA-enabled account too.
8. Exports report results to CSV.
9. The script is scheduler friendly.
10. It can be executed with certificate-based authentication (CBA) too.

For detailed script execution:https://o365reports.com/2024/05/21/export-exchange-online-mailbox-folder-statistics-using-powershell/

============================================================================================
#>

    Param (
    [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [string]$Organization,
        [string]$CertificateThumbprint,
        [string]$UserName,
        [string]$Password,
        [string]$MailboxUPN  ,
        [string]$MailBoxCSV ,
        [switch]$UserMailboxOnly,
        [switch]$SharedMailboxOnly,
        [string]$FolderPaths #Must folderpaths like(/Inbox,/Sent Items,/Inbox/SubFolder)
         )
    # Check for EXO module installation
    $Module = Get-Module ExchangeOnlineManagement -ListAvailable
    if ($Module.count -eq 0) { 
        Write-Host "Exchange Online PowerShell module is not available" -ForegroundColor yellow  
        $Confirm = Read-Host "Are you sure you want to install the module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-host "Installing Exchange Online PowerShell module"
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
        } 
        else {
            Write-Host "EXO module is required to connect Exchange Online. Please install module using Install-Module ExchangeOnlineManagement cmdlet." -ForegroundColor Yellow
            Exit
        }
    } 
    Write-Host "Connecting to Exchange Online..."
    # Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
    if (($UserName -ne "") -and ($Password -ne "")) {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
        Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
    }
    elseif ($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "") {
        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    
#Outputfile path
   $Location = Get-Location 
   $OutputCSV = "$($Location)\MailboxFolderStatisticsReports_$((Get-Date -format 'dd-MMM-yyyy-ddd hh-mm-ss').ToString()).csv"
  
#Function to FolderStatistics
 Function GetFolderStatistics($Folder){ 
        $FolderDetail = [PSCustomObject]@{
                        "Display Name" = $DisplayName
                        "UPN" =$MailboxUPN
                        "Folder Name" =$Folder.Name
                        "Folder Path" =$Folder.FolderPath
                        "Items In Folder" =$Folder.ItemsInFolder
                        "Folder Size" =$Folder.FolderSize.ToString().split("(") | Select-Object -Index 0
                        "Items In Folder And Subfolders" =$Folder.ItemsInFolderAndSubfolders 
                        "Folder And Subfolder Size" =$Folder.FolderAndSubfolderSize.ToString().split("(") | Select-Object -Index 0
                        "Deleted Items In Folder" =$Folder.DeletedItemsInFolder
                        "DeletedItems In Folder And Subfolders" =$Folder.DeletedItemsInFolderAndSubfolders
                        "Visible Items In Folder" =$Folder.VisibleItemsInFolder
                        "Hidden Items In Folder" =$Folder.HiddenItemsInFolder
                        "Mailbox Type" =$Mailboxtype
                        "Folder Type" =$Folder.FolderType
                        "Creation Time" =$Folder.CreationTime
                        "Last Modified Time" =$Folder.LastModifiedTime
                          }  
         $FolderDetail | Export-csv -Path $OutputCSV -Append -NoTypeInformation -Force 
        
}  
#Function to statistics for all folders 
Function ProcessAllMailboxFolders {
  Param([string]$MailboxUPN)
          $FolderDetails = Get-EXOMailboxFolderStatistics -Identity $MailboxUPN 
          if($FolderDetails){
          Foreach($Folder in $FolderDetails){
            GetFolderStatistics $Folder
          }
          Add-Content -Path "$OutputCSV" -Value "" 
  }
}
#Function to statistics for specific Folders
Function ProcessSpecificMailboxFolder {
    Param(
        [string]$MailboxUPN,
        [string]$FolderPaths
         )
        $Folders = $FolderPaths -split ',' |ForEach-Object { $_.Trim() }
        $FolderPathDetails = Get-EXOMailboxFolderStatistics -Identity $MailboxUPN | Where-Object { $_.FolderPath -in $Folders}
        if($FolderPathDetails){
           Foreach($Folder in $FolderPathDetails){
      
            GetFolderStatistics $Folder
            } 
            Add-Content -Path "$OutputCSV" -Value ""  
     }
 }
 #Function to get mailbox details
 Function Getmailbox{
   Param($MailboxUPN)
       $MailBoxInfo = Get-EXOMailbox -UserPrincipalName $MailboxUPN
       $DisplayName =$MailboxInfo.DisplayName
       $MailboxType = $MailboxInfo.RecipientTypeDetails
       Process-Mailbox -MailboxUPN $MailboxUPN
 }
 Function Process-Mailbox{   
   Param ($MailboxUPN)
       
        Write-Progress -Activity "Processed Mailbox count : $ProgressIndex" -Status "Currently Processing : $MailboxUPN"    
        if($Folderpaths -eq ""){
           ProcessAllMailboxFolders -MailboxUPN $MailboxUPN
        }
        else{
           ProcessSpecificMailboxFolder -MailboxUPN $MailboxUPN -FolderPaths $FolderPaths 
        }
}
 #Single user
 if ($MailboxUPN){ 
     $ProgressIndex =1 
     Getmailbox -MailboxUPN $MailboxUPN
     
 }

#Multiple CSV users
elseif($MailBoxCSV){ 
        $Mailboxes = Import-Csv -Path $MailBoxCSV
        $TotalMailboxes = $Mailboxes.Count
        $ProgressIndex = 0
        Foreach ($Mailbox in $Mailboxes) {
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
              
                         Write-Host "`nThe output file contains $($ProgressIndex) mailboxes and  $($ItemCounts) records." -ForegroundColor Cyan
         
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