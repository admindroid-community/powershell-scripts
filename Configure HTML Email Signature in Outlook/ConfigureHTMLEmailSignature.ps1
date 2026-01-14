<# -----------------------------------------------------------------------------------------------------------
Name           : How to Add HTML Signature in Outlook
Version        : 1.0
Website        : admindroid.com

~~~~~~~~~~~~~~~~~
Script Highlights: 
~~~~~~~~~~~~~~~~~

1. The script automatically verifies and installs the Exchange PowerShell module (if not installed already) upon your confirmation. 
2. Provides the option to create email signatures using HTML templates. 
3. Provides the option to use default or customized fields and templates. 
4. Allows to create an email signature for all mailboxes. 
5. Allows to filter and set up email signatures for user mailboxes alone. 
6. Allow to set up an email signature for bulk users. 
7. Exports signature deployment status to a CSV file. 
8. Supports certificate-based authentication (CBA) too. 

For detailed script execution: https://blog.admindroid.com/how-to-add-html-signature-in-outlook/

---------------------------------------------------------------------------------------------------------- #>

#Block for passing params
[CmdletBinding(DefaultParameterSetName = 'NoParams')]
param
(
  [Parameter()]
  [string]$Organization,
  [Parameter()]
  [string]$ClientId,
  [Parameter()]
  [string]$CertificateThumbprint,
  [Parameter()]
  [string]$UserPrincipalName,
  [Parameter()]
  [string]$Password,
  [Parameter(ParameterSetName = 'HTMLSignature_WithInbuiltHTML')]
  [switch]$AssignDefault_HTMLSignature,
  [Parameter(ParameterSetName = 'GetHTMLTemplateFromUser')]
  [switch]$AssignCustom_HTMLSignature,
  [Parameter()]
  [switch]$Enable_PostponeRoamingSignatureUntilLater,
  [Parameter()]
  [string]$UserListCsvPath,
  [Parameter()]
  [switch]$AllUsers,
  [Parameter()]
  [switch]$UserMailboxOnly,
  [Parameter(ParameterSetName = 'GetHTMLTemplateFromUser')]
  [string]$HTML_FilePath
)

#Check and Install Exchange Online Module
Function Installation-Module{
    $Module = (Get-Module ExchangeOnlineManagement -ListAvailable)
    if ($Module.count -eq 0)
    {
      Write-Host `n`Exchange Online PowerShell module is not available -ForegroundColor Red
      $Confirm = Read-Host `n`Are you sure you want to install module? [Y] Yes [N] No
      if ($Confirm -match "[yY]")
      {
        Write-Host "`n`Installing Exchange Online PowerShell module" -ForegroundColor Yellow
        Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
      }
      else
      {
        Write-Host "`n`EXO module is required to connect Exchange Online. Please install module using 'Install-Module ExchangeOnlineManagement' cmdlet." -ForegroundColor Yellow
        exit
      }
    }
    Import-Module ExchangeOnlineManagement
}

#Connecting to Exchangeonline
Function Connection-Module
{
    Write-Host "`n`Connecting to Exchange Online..."
    if ($UserPrincipalName -ne "" -and $Password -ne "")
    {
      $SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
      $UserCredential = New-Object System.Management.Automation.PSCredential ($UserPrincipalName,$SecurePassword)
      Connect-ExchangeOnline -Credential $UserCredential -ShowBanner:$false
    }
    elseif ($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
    {
      Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
    }
    else
    {
      Connect-ExchangeOnline -ShowBanner:$false
    }
}

#------------------------Block for getting the Confirmation from user to enable the PostponeRoamingSignaturesUntilLater parameter if not already enabled------------------------------------

Function Enable-PostponeRoamingSign
{
if (-not (Get-OrganizationConfig).PostponeRoamingSignaturesUntilLater)
{
  if (-not $Enable_PostponeRoamingSignatureUntilLater)
  {
    Write-Host "`n`To deploy signature for users, enable 'PostponeRoamingSignatureUntilLater' in the organization's settings." -ForegroundColor Yellow
    Write-Host "`n`1. Enable and continue.
    `n`2. Continue without enabling." -ForegroundColor Cyan
    $UserConfirmation = Read-Host "`n`Enter Your choice"
  }
  while ($true)
  {
    if ($Enable_PostponeRoamingSignatureUntilLater -or $UserConfirmation -eq 1)
    {
        Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true
        if($?)
        {
         Write-Host "`n`PostponeRoamingSignatureUntilLater parameter enabled.`n" -ForegroundColor Green
         break; 
        }
        else
        {
         Write-Host "`n`Failed to enable PostPoneRoamingSignaturesUntilLater: $($_.Exception.Message)" -ForegroundColor Red
         Exit;
        }
    }
    elseif ($UserConfirmation -eq 2)
    {
      Write-Host "`n`Proceeding without enabling 'PostPoneRoamingSignaturesUntilLater'. Signature will be added but not deployed to the mailboxes.`n" -ForegroundColor Yellow
      break;
    }
    else
    {
      Write-Host "`n`Enter the correct choice" -ForegroundColor Red
      $UserConfirmation = Read-Host
      continue;
    }
  }
}
else
{
  Write-Host "`n`'PostponeRoamingSignatureUntilLater' parameter already enabled`n"
  $Script:UserConfirmation = 1
}
}

#--------------------------------------------Function to preview the HTML Signature in your browser and get conformation to use that HTML template--------------------------------------------

function Preview-Signature ($FilePath)
{
  $FileExtension = [System.IO.Path]::GetExtension($FilePath)
  if ($FileExtension -eq ".html" -or $FileExtension -eq ".htm")
  {
    $HTMLFilePath = $FilePath
  }
  else
  {
    Write-Host "`n`The script will terminate as the file isn't in HTML format" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false
    exit
  }
  $Title = "Confirmation"
  $Question = "Do you want to preview the HTML Signature?"
  $Choices = "&Yes","&No"
  $Decision = $Host.UI.PromptForChoice($Title,$Question,$Choices,1)
  if ($Decision -eq 0) {
    Start-Process $HTMLFilePath
  }
  $UserChoice = Read-Host "`n`Are you sure to deploy the signature with this HTML template? [Y] Yes [N] No"
  while ($true)
  {
    if ($UserChoice -match "[yY]")
    {
      break
    }
    elseif ($UserChoice -match "[nN]")
    {
      Write-Host "`n`Exiting the script..."
      Disconnect-ExchangeOnline
      exit
    }
    else
    {
      Write-Host "`n`Invalid input. Enter the correct input" -ForegroundColor Red
      $UserChoice = Read-Host
    }
  }
  return
}

#---------------------------------------------function to get the users to assign an signature------------------------------------------------

function Get-UsersForAssignSignature
{
  if ($AllUsers)
  {
    $ImportUsersType = 1;
  }
  elseif ($UserListCsvPath)
  {
    $ImportUsersType = 2;
  }
  elseif ($UserMailboxOnly)
  {
    $ImportUsersType = 3;
  }
  else
  {
    Write-Host "    `n    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" -ForegroundColor Green
    Write-Host "                User-Based Email HTML Signature Deployment              " -ForegroundColor Yellow
    Write-Host "    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" -ForegroundColor Green
    write-host " "
    Write-Host "    1. Assign HTML signature to all type of mailboxes
    `n    2. Assign HTML signature to specific mailboxes (Import CSV)
    `n    3. Assign HTML signature to user mailboxes alone
    " -ForegroundColor Cyan
    $ImportUsersType = Read-Host "`n`Enter your choice"
  }
  while ($true)
  {
    if ($ImportUsersType -eq 1)
    {   
      $UsersCollection = Get-EXOMailbox -ResultSize Unlimited | Select-Object UserPrincipalName | Sort-Object -Property UserPrincipalName -Unique -ErrorAction Stop      
      $Output = $Output + " to $($UsersCollection.count) users in your tenant..."
      Write-Host $Output
    }
    elseif ($ImportUsersType -eq 2)
    {
      if ($UserListCsvPath)
      {
        $Path = $UserListCsvPath
      }
      else
      {
        $Path = Read-Host "`n`Enter the CSV file Path (Eg: C:\UsersCollection.csv)"
      }
      while ($true)
      {
        try
        {
          $Headers = (Import-Csv -Path $Path | Get-Member -MemberType NoteProperty).Name
          $UsersCollection = Import-Csv -Path $Path | Select-Object -Property UserPrincipalName | Sort-Object -Property UserPrincipalName -Unique -ErrorAction Stop
          if ('UserPrincipalName' -in $Headers -and $UsersCollection.count -ge 1)
          {
            $Output = $Output + "to $($UsersCollection.count) users in the CSV file..."
            Write-Host $Output
            break;
          }
          else
          {
            Write-Host "`n`The file is empty or does not contain the UserPrincipalName column." -ForegroundColor Red
            $Path = Read-Host "`n`Enter correct file path"
            continue
          }
        }
        catch
        {
          Write-Host "`n`Invalid path provided." -ForegroundColor Red
          $Path = Read-Host "`n`Enter the correct path (Eg: C:\UsersCollection.csv)"
        }
      }
    }
    elseif($ImportUsersType -eq 3)
    {
      $UsersCollection = Get-EXOMailbox -RecipientTypeDetails UserMailBox -ResultSize Unlimited | Select-Object UserPrincipalName | Sort-Object -Property UserPrincipalName -Unique -ErrorAction Stop
      $Output = $Output + " to $($UsersCollection.count) users in your tenant..."
      Write-Host $Output
    }
    else
    {
      Write-Host "`n`Enter the valid input : " -ForegroundColor Red
      $ImportUsersType = Read-Host
      continue;
    }
    return $UsersCollection
  }

}

#-------------------------------------------------------------Function to get the content of the HTML with the specified Path----------------------------------------------------------

function Get-HTMLContent
{
  if (-not $HTML_FilePath)
  {
    $Path = Read-Host "`n`Enter the HTML file Path (eg: C:\SigantureTemplate.html)"
  }
  else
  {
    $Path = $HTML_FilePath
  }
  while ($true)
  {
    try
    {
      $HTMLcontent = Get-Content -Path $Path -Raw -ErrorAction Stop
      if ($HTMLcontent.Length -eq 0)
      {
        Write-Host '`n`Provided file is empty.' -ForegroundColor Red
        $Path = Read-Host '`n`Enter the valid HTML file Path'
        continue
      }
      break;
    }
    catch
    {
      Write-Host "`n`Invalid Path. Enter the correct Path (eg: C:\SigantureTemplate.html)" -ForegroundColor Red
      $Path = Read-Host
    }
  }
  Preview-Signature ($Path)
  return [string]$HTMLcontent
}

#-------------------------------------------------------------Function to get the user required fields for custom template----------------------------------------------------------

function Get-RequiredFieldsFromUser
{
  Write-Host @"
    `n`You can set up signature using the below fields.
    `n`1. Display Name
    `n`2. Email Address     
    `n`3. Mobile Phone 
    `n`4. Business Phone
    `n`5. Fax     
    `n`6. Department 
    `n`7. Title 
    `n`8. Office 
    `n`9. Address (Street Address, City, State or Province, Zip Code, Country or Region)
    `n`10. StreetAddress
    `n`11. City
    `n`12. StateOrProvince
    `n`13. PostalCode
    `n`14. CountryOrRegion
"@ -ForegroundColor Yellow

  $FieldNumbers = Read-Host "`n`Enter the field numbers to update in the HTML signature (for example: 1,5,6,14)"
  $Numbers = $FieldNumbers -split ','
  return $Numbers
}

#-----------------------------------------------------------------Function to get the user address ----------------------------------------------------------------------

function Generate-UserAddress ($UserDetails)
{
  if ($UserDetails.StreetAddress)
  {
    $Address += $UserDetails.StreetAddress + ", "
  }
  if ($UserDetails.City)
  {
    $Address += $UserDetails.City + ", "
  }
  if ($UserDetails.StateOrProvince)
  {
    $Address += $UserDetails.StateOrProvince + ",<br>"
  }
  if ($UserDetails.PostalCode)
  {
    $Address += $UserDetails.PostalCode + ", "
  }
  if ($UserDetails.CountryOrRegion)
  {
    $Address += $UserDetails.CountryOrRegion + "."
  }
  return $Address
}

#---------------------------------------------------------------Function to generate and Assign inbuild HTML Signature to the users-------------------------------------------------------------------

function Deploy-InbuiltHTMLSignature
{
  $Filepath = Join-Path $PSScriptRoot "Build-InTemplate.html"
  try
  {
    $DefaultHTML = Get-Content $Filepath -ErrorAction Stop
  }
  catch
  {
    Write-Host "`n`Inbuilt HTMLFile is not available in the current folder" -ForegroundColor Red
    Write-Host "`n`Script execution stopped..."
    Disconnect-ExchangeOnline
    exit
  }
  Preview-Signature $Filepath
  Write-Host "`n`Enter values for the upcoming fields to include in the signature, or press Enter to skip." -ForegroundColor Yellow
  $LogoLink = Read-Host "`n`Enter your company's logo link ( https:// )"
  if ($LogoLink -eq "")
  {
    $DefaultHTML = $DefaultHTML.Replace('<img width="125px" src="https://cdn0.iconfinder.com/data/icons/logos-microsoft-office-365/128/Microsoft_Office-07-1024.png">','')
  }
  else
  {
    $DefaultHTML = $DefaultHTML.Replace('https://cdn0.iconfinder.com/data/icons/logos-microsoft-office-365/128/Microsoft_Office-07-1024.png',$LogoLink)
  }
  $WebLink = Read-Host "`n`Enter your company's website link ( https:// )"
  if ($WebLink -eq "")
  {
    $DefaultHTML = $DefaultHTML.Replace('<a href="%%WebLink%%" target="_blank">%%WebLink%%</a>','')
  }
  else
  {
    $DefaultHTML = $DefaultHTML.Replace('%%WebLink%%',$WebLink)
  }
  $Facebook = Read-Host "`n`Enter your company's Facebook link ( https:// )"
  if ($Facebook -eq "")
  {
    $DefaultHTML = $DefaultHTML.Replace('<a href="%%Facebook%%" target="_blank"><img style="height:22px; width:26px;" src="https://cdn0.iconfinder.com/data/icons/social-flat-rounded-rects/512/facebook-1024.png"></a>','')
  }
  else
  {
    $DefaultHTML = $DefaultHTML.Replace('%%Facebook%%',$FaceBook)
  }
  $Twitter = Read-Host "`n`Enter your company's X link ( https:// )"
  if ($Twitter -eq "")
  {
    $DefaultHTML = $DefaultHTML.Replace('<a href="%%Twitter%%" target="_blank"><img style="height:22px; width:26px;" src="https://cdn2.iconfinder.com/data/icons/threads-by-instagram/24/x-logo-twitter-new-brand-contained-1024.png"></a>','')
  }
  else
  {
    $DefaultHTML = $DefaultHTML.Replace('%%Twitter%%',$Twitter)
  }
  $YouTube = Read-Host "`n`Enter your company's YouTube link ( https:// )"
  if ($YouTube -eq "")
  {
    $DefaultHTML = $DefaultHTML.Replace('<a href="%%YouTube%%" target="_blank"><img style="height:22px; width:26px;" src="https://cdn0.iconfinder.com/data/icons/web-social-and-folder-icons/512/YouTube.png"></a>','')
  }
  else
  {
    $DefaultHTML = $DefaultHTML.Replace('%%YouTube%%',$YouTube)
  }
  $LinkedIn = Read-Host "`n`Enter your company's LinkedIn link ( https:// )"
  if ($LinkedIn -eq "")
  {
    $DefaultHTML = $DefaultHTML.Replace('<a href="%%LinkedIn%%" target="_blank"><img style="height:22px; width:26px;" src="https://cdn2.iconfinder.com/data/icons/social-media-2285/512/1_Linkedin_unofficial_colored_svg-1024.png"></a>','')
  }
  else
  {
    $DefaultHTML = $DefaultHTML.Replace('%%LinkedIn%%',$LinkedIn)
  }
  $DisCord = Read-Host "`n`Enter your company's Discord link ( https:// )"
  if ($DisCord -eq "")
  {
    $DefaultHTML = $DefaultHTML.Replace('<a href="%%Discord%%" target="_blank"><img style="height:22px; width:26px;" src="https://cdn2.iconfinder.com/data/icons/gaming-platforms-squircle/250/discord_squircle-1024.png"></a>','')
  }
  else
  {
    $DefaultHTML = $DefaultHTML.Replace('%%Discord%%',$DisCord)
  }
  $DefaultHTML = [string]$DefaultHTML
  $UsersCollection = Get-UsersForAssignSignature
  $Count = 1
  $TotalCount = $UsersCollection.count
  $CurrentTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
  $SignatureLog_FilePath = "$(Get-Location)\SignatureDeployment_Details_$CurrentTime.csv"
  foreach ($User in $UsersCollection)
  {
    try
    {
      $UserDetails = Get-User -Identity $User.UserPrincipalName -ErrorAction Stop
      $Address = Generate-UserAddress ($UserDetails)
      $UserHTMLSignature = $DefaultHTML -replace "%%DisplayName%%",$UserDetails.DisplayName -replace "%%Title%%",$UserDetails.Title -replace "%%Email%%",$UserDetails.UserPrincipalName -replace "%%MobilePhone%%",$UserDetails.MobilePhone -replace "%%BusinessPhone%%",$UserDetails.Phone -replace "%%CompanyName%%",$UserDetails.Office -replace "%%Address%%",$Address
      Set-MailboxMessageConfiguration -Identity $UserDetails.UserPrincipalName -SignatureHTML $UserHTMLSignature -AutoAddSignature $true -AutoAddSignatureOnMobile $true -AutoAddSignatureOnReply $true -ErrorAction Stop
      $DeploymentStatus = "Successful"
    }
    catch
    {
      $DeploymentStatus = "Failed"
      $ErrorMessage = $_.Exception.Message
    }
    $AuditData = [pscustomobject]@{ UserPrincipalName = $User.UserPrincipalName; DeploymentStatus = $DeploymentStatus; Error = if ($ErrorMessage) { $ErrorMessage } else { "-" }}
    $AuditData | Export-Csv -Path $SignatureLog_FilePath -NoTypeInformation -Append
    Write-Progress -Activity "Deploying Signature : $($Count) of $($TotalCount) | Current User : $($User.UserPrincipalName)"
    $Count++
  }
  Disconnect_ExchangeOnline_Safely
}

#---------------------------------------------------------------Function to generate and Assign custom HTML Signature to the users-------------------------------------------------------------------

function Deploy-CustomHTMLSignature
{
  $HTMLSignature = Get-HTMLContent
  $Numbers = Get-RequiredFieldsFromUser
  $AddressUsed = $false
  $UsersFields = @{ '1' = 'DisplayName'; '2' = 'EmailAddress'; '3' = 'MobilePhone'; '4' = 'BussinessPhone'; '5' = 'FaxNumber'; '6' = 'Department'; '7' = 'Title'; '8' = 'Office'; '9' = 'Address'; '10' = 'StreetAddress'; '11' = 'City'; '12' = 'StateOrProvince'; '13' = 'PostalCode'; '14' = 'CountryOrRegion' }
  Write-Host "`n`We need the current values of the following fields to update them with user-specific details." -ForegroundColor Yellow
  foreach ($Number in $Numbers) {
    if ($UsersFields[$Number])
    {
      Write-Host "`n`Enter the $($UsersFields[$Number]) as it appears in the provided HTML:" -ForegroundColor Cyan
      while ($true)
      {
        $ValueInHTML = Read-Host
        if ($ValueInHTML -ne "")
        {
          break
        }
        else
        {
          Write-Host "`n`Enter the valid information" -ForegroundColor Red
        }
      }
      if ($Number -eq '9')
      {
        $AddressUsed = $true
      }
      if ($ValueInHTML -match "(&|&amp;)")
      {
        $ValueInHTML = $ValueInHTML -replace "(&|&amp;)","&amp;"
      }
      $HTMLSignature = $HTMLSignature -replace [regex]::Escape($ValueInHTML),"%%$($UsersFields[$Number])%%"      
    }
  }
  $UsersCollection = Get-UsersForAssignSignature
  $Count = 1
  $TotalCount = $UsersCollection.count
  $CurrentTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
  $SignatureLog_FilePath = "$(Get-Location)\SignatureDeployment_Details_$CurrentTime.csv"
  foreach ($User in $UsersCollection)
  {
    try
    {
      $UserDetails = Get-User -Identity $User.UserPrincipalName -ErrorAction Stop
      if ($AddressUsed) {
        $Address = Generate-UserAddress ($UserDetails)
      }
      $UserSignature = $HTMLSignature.Replace('%%DisplayName%%',$UserDetails.DisplayName).Replace('%%EmailAddress%%',$UserDetails.UserPrincipalName).Replace('%%MobilePhone%%',$UserDetails.MobilePhone).Replace('%%BussinessPhone%%',$UserDetails.Phone).Replace('%%FaxNumber%%',$UserDetails.Fax).Replace('%%Department%%',$UserDetails.Department).Replace('%%Title%%',$UserDetails.Title).Replace('%%Office%%',$UserDetails.Office).Replace('%%Address%%',$Address).Replace('%%StreetAddress%%',$UserDetails.StreetAddress).Replace('%%City%%',$UserDetails.City).Replace('%%StateOrProvince%%',$UserDetails.StateOrProvince).Replace('%%PostalCode%%',$UserDetails.PostalCode).Replace('%%CountryOrRegion%%',$UserDetails.CountryOrRegion)
      Set-MailboxMessageConfiguration -Identity $UserDetails.UserPrincipalName -SignatureHTML $UserSignature -AutoAddSignature $true -AutoAddSignatureOnMobile $true -AutoAddSignatureOnReply $true -ErrorAction Stop
      $DeploymentStatus = "Successful"
    }
    catch
    {
      $DeploymentStatus = "Failed"
      $ErrorMessage = $_.Exception.Message
    }
    $AuditData = [pscustomobject]@{ UserPrincipalName = $User.UserPrincipalName; DeploymentStatus = $DeploymentStatus; Error = if ($ErrorMessage) { $ErrorMessage } else { "-" } }
    $AuditData | Export-Csv -Path $SignatureLog_FilePath -NoTypeInformation -Append
    Write-Progress -Activity "Deploying Signature : $($Count) of $($TotalCount) | Current User : $($User.UserPrincipalName)"
    $Count++
  }
  Disconnect_ExchangeOnline_Safely
}

#-------------------function to verify the status of Signature Assign to user and process the auditlog csv file and disconnect the exchangeonline module-------------------------

function Disconnect_ExchangeOnline_Safely
{
  if ($Script:UserConfirmation -ne 1)
  {
    Write-Host "`n`The HTML signature has been added and can be deployed when 'PostponeRoamingSignatureUntilLater' enabled."
  }

  if((Test-Path -Path $SignatureLog_FilePath) -eq "True")
  {
    Write-Host "`n`Script execution completed." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "`n`The signature deployment status report available in: " -NoNewline -ForegroundColor Yellow
    Write-Host $SignatureLog_FilePath
    $Prompt = New-Object -ComObject wscript.shell   
    $UserInput = $Prompt.popup("Do you want to open status report?",0,"Open Status Report",4)   
    If ($UserInput -eq 6)   
    {   
    Invoke-Item "$SignatureLog_FilePath"   
    }
  }
  else
  {
    Write-Host "No logs found"
  }
  Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
  Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green `n`n
  exit
}

#---------------------------------------Block to call the function as per the user Choice----------------------------
Installation-Module
Connection-Module
Enable-PostponeRoamingSign

$Output = ""
if (-not ($AssignDefault_HTMLSignature -or $AssignCustom_HTMLSignature))
{
  Write-Host "    =================================================================" -ForegroundColor Green
  Write-Host "                Mailbox HTML Signature configuration         " -ForegroundColor Yellow
  Write-Host "    =================================================================" -ForegroundColor Green
  Write-Host " "
  Write-Host "    1. Create an HTML Signature Using an In-built HTML Template
  `n    2. Create an HTML Signature with a Custome Template 
  " -ForegroundColor Cyan
  $UserChoice = Read-Host "`n`Enter your choice"
}
while ($true)
{
  if ($AssignDefault_HTMLSignature -or $UserChoice -eq 1)
  {
    $UserChoice = 0
    $Output = "`n`Adding In-build HTML signature"
    Deploy-InbuiltHTMLSignature
  }
  elseif ($AssignCustom_HTMLSignature -or $UserChoice -eq 2)
  {
    $UserChoice = 0
    $Output = "`n`Adding Custom HTML signature"
    Deploy-CustomHTMLSignature
  }
  else
  {
    Write-Host "`n`Invalid input. Enter the valid choice" -ForegroundColor Red
    $UserChoice = Read-Host
    continue;
  }
}
