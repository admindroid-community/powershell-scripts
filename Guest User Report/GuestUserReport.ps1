#Accept input parameter
Param
(
    [Parameter(Mandatory = $false)]
    [int]$StaleGuests,
    [int]$RecentlyCreatedGuests,
    [string]$UserName,
    [string]$Password
    
)

#Check for AzureAD module
$Module=Get-Module -Name AzureAD -ListAvailable  
if($Module.count -eq 0) 
{ 
 Write-Host AzureAD module is not available  -ForegroundColor yellow  
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
 if($Confirm -match "[yY]") 
 { 
  Install-Module AzureAD
  Import-Module AzureAD
 } 
 else 
 { 
  Write-Host AzureAD module is required to connect AzureAD.Please install module using Install-Module AzureAD cmdlet. 
  Exit
 }
} 
 
Write-Host Connecting Azure AD... -ForegroundColor Yellow
#Storing credential in script for scheduling purpose/ Passing credential as parameter  
if(($UserName -ne "") -and ($Password -ne ""))  
{  
 $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force  
 $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword  
 Connect-AzureAD -Credential $credential | Out-Null
}  
else  
{  
 Connect-AzureAD | Out-Null  
} 

$Result=""   
$Results=@()  
$GuestCount=0
$PrintedGuests=0

#Output file declaration 
$ExportCSV=".\GuestUserReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
Write-Host `nExporting report... 
#Getting guest users
Get-AzureADUser -All $true -Filter "UserType eq 'Guest'" | foreach {
 $DisplayName=$_.DisplayName
 $Upn=$_.UserPrincipalName
 $GuestCount++
 Write-Progress -Activity "`n     Processed mailbox count: $GuestCount "`n"  Currently Processing: $DisplayName"
 $GetCreationTime=$_.ExtensionProperty
 $CreationTime=$GetCreationTime.createdDateTime 
 $AccountAge= (New-TimeSpan -Start $CreationTime).Days

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

 $ObjectId=$_.ObjectId
 $Mail=$_.Mail
 $Company=$_.CompanyName
 if($Company -eq $null)
 {
  $Company="-"
 }
 $CreationType=$_.CreationType
 $InvitationAccepted=$_.UserState

 #Getting guest user's group membership
 $Groups=(Get-AzureADUserMembership -ObjectId $ObjectId).DisplayName
 #$Groups
 $GroupMembership=""
 foreach($Group in $Groups)
 {
  #$Group
  if($GroupMembership -ne "")
  {
   $GroupMembership=$GroupMembership+","
  }
  $GroupMembership=$GroupMembership+$Group
 }
 if($GroupMembership -eq "")
 {
  $GroupMembership="-"
 }
 

 #Export result to CSV file 
 $PrintedGuests++
 $Result=@{'UserPrincipalName'=$upn;'DisplayName'=$DisplayName;'EmailAddress'=$Mail;'Company'=$Company;'CreationTime'=$CreationTime;'AccountAge(days)'=$AccountAge;'InvitationAccepted'=$InvitationAccepted;'CreationType'=$CreationType; 'GroupMembership'=$GroupMembership} 
 $Output= New-Object PSObject -Property $Result 
 $Output | Select-Object DisplayName,UserPrincipalName,Company,EmailAddress,CreationTime,AccountAge'(Days)',CreationType,InvitationAccepted,GroupMembership | Export-Csv -Path $ExportCSV -Notype -Append
}

 #Open output file after execution 
 Write-Host `nScript executed successfully
 if((Test-Path -Path $ExportCSV) -eq "True")
 {
  Write-Host "Detailed report available in: $ExportCSV" -ForegroundColor Green
  Write-Host `nThe Output file contains $PrintedGuests guest users.
  $Prompt = New-Object -ComObject wscript.shell  
  $UserInput = $Prompt.popup("Do you want to open output file?",`  
 0,"Open Output File",4)  
  If ($UserInput -eq 6)  
  {  
   Invoke-Item "$ExportCSV"  
  } 
 }
 Else
 {
  Write-Host No guest user found
 }