<#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Name: Automate Email Signature Setup in Outlook Using PowerShell
Version: 2.0
Website: o365reports.com

~~~~~~~~~~~~~~~~~
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script automatically verifies and installs the Exchange PowerShell module (if not installed already) upon your confirmation.
2. Automates email signature deployment and makes it scheduler friendly. 
3. Provides the option to create default or custom text signatures. 
4. Provides the option to configure default or custom signatures using HTML templates. 
5. Automates email signature setup for all mailboxes or bulk mailboxes. 
6. Automates email signature configuration for user mailboxes alone. 
7. Exports signature deployment status to a CSV file. 
8. Supports certificate-based authentication (CBA) too.

For detailed script execution: https://o365reports.com/2024/07/10/automate-email-signature-setup-in-outlook-using-powershell/


Change Log:
~~~~~~~~~~~

v1.0 (July 10, 2024)- Script created
V2.0 (Jan 18, 2025)- Error handling added to enabling PostponeRoamingSignatureUntilLater param.
------------------------------------------------------------------------------------------------------------#>

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
  [Parameter()]
  [switch]$TaskScheduler
)

#--------------------------------------Block For Module Availability Verification and Installation--------------------------------------------

$Module = (Get-Module ExchangeOnlineManagement -ListAvailable)
if ($Module.count -eq 0)
{
  Write-Host `n`Exchange Online PowerShell module is not available
  $Confirm = Read-Host `n`Are you sure you want to install module? [Y] Yes [N] No
  if ($Confirm -match "[yY]")
  {
    Write-Host "`n`Installing Exchange Online PowerShell module" -ForegroundColor Red
    Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
    Import-Module ExchangeOnlineManagement
  }
  else
  {
    Write-Host `n`EXO module is required to connect Exchange Online. Please install module using Install-Module ExchangeOnlineManagement cmdlet. -ForegroundColor Yellow
    exit
  }
}

#-------------------------------------------------------------------------Block For Connecting to Exchangeonline ---------------------------------------------------------------------------
if($TaskScheduler)
{
    Write-Host "Setup email signature in Outlook - Scheduled task started" -ForegroundColor Cyan
}
Write-Host `n`Connecting to Exchange Online... -ForegroundColor Green
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
elseif (-not $TaskScheduler)
{
  Connect-ExchangeOnline -ShowBanner:$false
}
else
{
  Write-Host "`n`The script will exit as the Required parameters are not included." -ForegroundColor Red
  exit
}
$InputsFolderPath = Join-Path $PSScriptRoot -ChildPath "StoredUserInputs.csv"
if (-not $TaskScheduler)
{
  if ((Test-Path -Path $InputsFolderPath))
  {
    Remove-Item -Path $InputsFolderPath
  }
}
else
{
  $Inputs = Import-Csv -Path $InputsFolderPath
  $Index = 0
}

#------------------------Block for getting the Confirmation from user to enable the PostponeRoamingSignaturesUntilLater parameter if not already enabled------------------------------------

if ((-not (Get-OrganizationConfig).PostponeRoamingSignaturesUntilLater) -and (-not $TaskScheduler))
{
  if (-not $Enable_PostponeRoamingSignatureUntilLater)
  {
    Write-Host "`n`To add a signature for users, first enable 'PostponeRoamingSignatureUntilLater' in the organization's settings." -ForegroundColor Yellow
    Write-Host "`n`1. Enable it." -ForegroundColor Cyan
    Write-Host "`n`2. Continue without enabling." -ForegroundColor Cyan
    $UserConfirmation = Read-Host "`n`Enter Your choice"
  }
  while ($true)
  {
    if ($Enable_PostponeRoamingSignatureUntilLater -or $UserConfirmation -eq 1)
    {
     Set-OrganizationConfig -PostponeRoamingSignaturesUntilLater $true
     if($?)
     {
      Write-Host "`n`PostponeRoamingSignatureUntilLater parameter enabled" -ForegroundColor Green
      break; 
     }
     else
     {
      Write-Host "Error occurred. Unable to enable PostPoneRoamingSignaturesUntilLater.Please try again" -ForegroundColor Red
      Exit;
     } 
    }
    elseif ($UserConfirmation -eq 2)
    {
      Write-Host '`n`Without Enabling it, Signature can be added but not deployed to Users MailBox' -ForegroundColor Yellow
      break;
    }
    else
    {
      Write-Host "`n`Enter the correct input" -ForegroundColor Red
      $UserConfirmation = Read-Host
      continue;
    }
  }
}
elseif (-not $TaskScheduler)
{
  Write-Host "`n`PostponeRoamingSignatureUntilLater parameter already enabled" -ForegroundColor Green
  $UserConfirmation = 1
}

#--------------------------------------------Function to preview the HTML Signaute in your browser and get conformation to use that HTML template--------------------------------------------

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
  Write-Host @"
`n`Are you sure to deploy the signature with this template? [Y] Yes [N] No
"@ -ForegroundColor Cyan
  $UserChoice = Read-Host "`n`Enter your choice"
  while ($true)
  {
    if ($UserChoice -match "[Y]")
    {
      break
    }
    elseif ($UserChoice -match "[N]")
    {
      Write-Host "`n`Exiting the script..."
      Disconnect-ExchangeOnline
      exit
    }
    else
    {
      Write-Host "`n`Enter the correct input" -ForegroundColor Red
      $UserChoice = Read-Host
    }
  }
  return
}

#---------------------------------------------function to get the users to assign an signature------------------------------------------------

function Get-UsersForAssignSignature
{
  if (-not $TaskScheduler)
  {
    Write-Host "`n`User-Based Email Signature Deployment:" -ForegroundColor Yellow
    Write-Host @"
                `n`1. Assign signature to all mailboxs.
                `n`2. Assign signature to specific mailboxs (Import CSV).
                `n`3. Assign signature to user mailboxs. 
"@ -ForegroundColor Cyan
    $ImportUsersType = Read-Host "`n`Enter your choice"
  }
  else
  {
    $ImportUsersType = $Inputs[$Index++].Input
  }
  while ($true)
  {
    if (($ImportUsersType -eq 1 -or $ImportUsersType -eq 2 -or $ImportUsersType -eq 3) -and (-not $TaskScheduler))
    {
      $ImportUsersType | Select-Object @{ Name = 'Input'; Expression = { $_ } } | Export-Csv -Path $InputsFolderPath -NoTypeInformation -Append
    }
    if ($ImportUsersType -eq 1)
    {
      $UsersCollection = Get-EXOMailbox -ResultSize Unlimited | Select-Object UserPrincipalName | Sort-Object -Property UserPrincipalName -Unique -ErrorAction Stop
      $Output = $Output + "to $($UsersCollection.count) Users in the tenant..."
      Write-Host $Output
    }
    elseif ($ImportUsersType -eq 2)
    {
      if (-not $TaskScheduler)
      {
        Write-Host "`n`Enter the csv file Path without quotation (Eg: C:\UsersCollection.csv)" -ForegroundColor Cyan
        $Path = Read-Host
      }
      else
      {
        $Path = $Inputs[$Index++].Input
      }
      while ($true)
      {
        try
        {
          $Headers = (Import-Csv -Path $Path | Get-Member -MemberType NoteProperty).Name
          $UsersCollection = Import-Csv -Path $Path | Select-Object -Property UserPrincipalName | Sort-Object -Property UserPrincipalName -Unique -ErrorAction Stop
          if ('UserPrincipalName' -in $Headers -and $UsersCollection.count -ge 1)
          {
            if (-not $TaskScheduler)
            {
              $Path | Select-Object @{ Name = 'Input'; Expression = { $_ } } | Export-Csv -Path $InputsFolderPath -NoTypeInformation -Append
            }
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
          Write-Host "`n`You have entered an invalid Path. Enter the correct Path (Eg: C:\UsersCollection.csv)" -ForegroundColor Red
          $Path = Read-Host
        }
      }
    }
    elseif ($ImportUsersType -eq 3)
    {
      $UsersCollection = Get-EXOMailbox -RecipientTypeDetails UserMailBox -ResultSize Unlimited | Select-Object UserPrincipalName | Sort-Object -Property UserPrincipalName -Unique -ErrorAction Stop
      $Output = $Output + "to $($UsersCollection.count) Users in the tenant..."
      Write-Host $Output
    }
    else
    {
      Write-Host "`n`Enter the correct input : " -ForegroundColor Red
      $ImportUsersType = Read-Host
      continue;
    }
    return $UsersCollection
  }

}

#-------------------------------------------------------------Function to get the content of the HTML with the specified Path----------------------------------------------------------

function Get-HTMLContent
{
  Write-Host "`n`Enter the HTML file Path without quotation (eg: C:\HTMLfile.html)" -ForegroundColor Cyan
  $Path = Read-Host
  while ($true)
  {
    try
    {
      $HTMLcontent = Get-Content -Path $Path -Raw -ErrorAction Stop
      if ($HTMLcontent.Length -eq 0)
      {
        Write-Host '`n`You have Provided an empty file' -ForegroundColor Red
        $Path = Read-Host '`n`Enter the valid HTML file Path'
        continue
      }
      break;
    }
    catch
    {
      Write-Host "`n`You have entered an invalid Path. Enter the correct Path (eg: C:\HTMLfile.html)" -ForegroundColor Red
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
    `n`You can set up signature using the below fields. If a field has a value in the admin center, that value will be included in the signature.
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
  if ($UserChoice -eq 3)
  {
    Write-Host "`n`Enter the field numbers from the list above that you would like to include in the text signature.(Eg: 1,5,6,14)" -ForegroundColor Cyan
  }
  else
  {
    Write-Host "`n`Provide the field numbers that need to be change in your HTML signature with users information present in admin center (Eg: 1,5,6,14)" -ForegroundColor Cyan
  }
  $FieldNumbers = Read-Host "`n`Enter Your choice"
  if (-not $TaskScheduler -and $UserChoice -eq 3)
  {
    $FieldNumbers | Select-Object @{ Name = 'Input'; Expression = { $_ } } | Export-Csv -Path $InputsFolderPath -NoTypeInformation -Append
  }
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

#---------------------------------------------------------------Function to generate and Assign default text Signature to the users-------------------------------------------------------------------

function Deploy-DefaultTextSignature
{
  $UsersCollection = Get-UsersForAssignSignature
  $Count = 1
  $TotalCount = $UsersCollection.count
  $CurrentTime = Get-Date -Format "yyyyMMdd_HHmmss"
  $SignatureLog_FilePath = Join-Path $PSScriptRoot -ChildPath "$CurrentTime-SignatureDeployment_Details.csv"
  foreach ($User in $UsersCollection)
  {
    try {
      $UserDetails = Get-User -Identity $User.UserPrincipalName -ErrorAction Stop
      if ($UserDetails.DisplayName -ne "")
      {
        $UserTextSignature += "$($UserDetails.DisplayName)<br>"
      }
      if ($UserDetails.UserPrincipalName -ne "")
      {
        $UserTextSignature += "$($UserDetails.UserPrincipalName)<br>"
      }
      if ($UserDetails.Title -ne "")
      {
        $UserTextSignature += "$($UserDetails.Title)<br>"
      }
      if ($UserDetails.MobilePhone -ne "")
      {
        $UserTextSignature += "$($UserDetails.MobilePhone)<br>"
      }
      if ($UserDetails.Phone)
      {
        $UserTextSignature += "$($UserDetails.Phone)<br>"
      }
      $UserTextSignature += Generate-UserAddress ($UserDetails)
      Set-MailboxMessageConfiguration -Identity $UserDetails.UserPrincipalName -SignatureHTML $UserTextSignature -AutoAddSignature $true -AutoAddSignatureOnMobile $true -AutoAddSignatureOnReply $true -ErrorAction Stop
      $DeploymentStatus = "Successful"
    }
    catch
    {
      $ErrorMessage = $_.Exception.Message
      $DeploymentStatus = "Unsuccessful : $ErrorMessage "
      Write-Host "`n`Signature deployment for $($User.UserPrincipalName) is Unsucessful -Error Occured" -ForegroundColor Red
    }
    $AuditData = [pscustomobject]@{ UserPrincipalName = $User.UserPrincipalName; DeploymentStatus = $DeploymentStatus }
    $AuditData | Export-Csv -Path $SignatureLog_FilePath -NoTypeInformation -Append
    Write-Progress -Activity "Deploying Signature : $($Count) of $($TotalCount) | Current User : $($User.UserPrincipalName)"
    $Count++
    $UserTextSignature = ""
  }
  Disconnect_ExchangeOnline_Safely
}

#---------------------------------------------------------------Function to generate and Assign inbuild HTML Signature to the users-------------------------------------------------------------------

function Deploy-InbuiltHTMLSignature
{
  if (-not $TaskScheduler)
  {
    $Filepath = Join-Path $PSScriptRoot "Build-InTemplate.html"
    try
    {
      $DefaultHTML = Get-Content $Filepath -ErrorAction Stop
    }
    catch
    {
      Write-Host "`n`Inbuilt HTMLFile is not available in the current directory" -ForegroundColor Red
      Write-Host "`n`Exiting the Script..."
      Disconnect-ExchangeOnline
      exit
    }
    Preview-Signature $Filepath
    Write-Host "`n`You can include upcoming fields in the signature by entering a value or simply pressing Enter to skip." -ForegroundColor Yellow
    $LogoLink = Read-Host "`n`Enter your company's logo link ( https:// )"
    if ($LogoLink -eq "")
    {
      $DefaultHTML = $DefaultHTML.Replace('<img width="125px" src="https://cdn0.iconfinder.com/data/icons/logos-microsoft-office-365/128/Microsoft_Office-07-1024.png">','')
    }
    else
    {
      $DefaultHTML = $DefaultHTML.Replace('https://cdn0.iconfinder.com/data/icons/logos-microsoft-office-365/128/Microsoft_Office-07-1024.png',$LogoLink)
    }
    $WebLink = Read-Host "`n`Enter your company's website link   ( https:// )"
    if ($WebLink -eq "")
    {
      $DefaultHTML = $DefaultHTML.Replace('<a href="%%WebLink%%" target="_blank">%%WebLink%%</a>','')
    }
    else
    {
      $DefaultHTML = $DefaultHTML.Replace('%%WebLink%%',$WebLink)
    }
    $FaceBook = Read-Host "`n`Enter your company's FaceBook link ( https:// )"
    if ($FaceBook -eq "")
    {
      $DefaultHTML = $DefaultHTML.Replace('<a href="%%FaceBook%%" target="_blank"><img style="height:22px; width:26px;" src="https://cdn0.iconfinder.com/data/icons/social-flat-rounded-rects/512/facebook-1024.png"></a>','')
    }
    else
    {
      $DefaultHTML = $DefaultHTML.Replace('%%FaceBook%%',$FaceBook)
    }
    $Twitter = Read-Host "`n`Enter your company's Twitter link ( https:// )"
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
    $LinkedIN = Read-Host "`n`Enter your company's LinkedIn link ( https:// )"
    if ($LinkedIN -eq "")
    {
      $DefaultHTML = $DefaultHTML.Replace('<a href="%%LinkedIN%%" target="_blank"><img style="height:22px; width:26px;" src="https://cdn2.iconfinder.com/data/icons/social-media-2285/512/1_Linkedin_unofficial_colored_svg-1024.png"></a>','')
    }
    else
    {
      $DefaultHTML = $DefaultHTML.Replace('%%LinkedIN%%',$LinkedIN)
    }
    $DisCord = Read-Host "`n`Enter your company's DisCord link ( https:// )"
    if ($DisCord -eq "")
    {
      $DefaultHTML = $DefaultHTML.Replace('<a href="%%Discord%%" target="_blank"><img style="height:22px; width:26px;" src="https://cdn2.iconfinder.com/data/icons/gaming-platforms-squircle/250/discord_squircle-1024.png"></a>','')
    }
    else
    {
      $DefaultHTML = $DefaultHTML.Replace('%%DisCord%%',$DisCord)
    }
    $DefaultHTML = [string]$DefaultHTML
    $DefaultHTML | Select-Object @{ Name = 'Input'; Expression = { $_ } } | Export-Csv -Path $InputsFolderPath -NoTypeInformation -Append
  }
  else
  {
    $DefaultHTML = $Inputs[$Index++].Input
  }
  $UsersCollection = Get-UsersForAssignSignature
  $Count = 1
  $TotalCount = $UsersCollection.count
  $CurrentTime = Get-Date -Format "yyyyMMdd_HHmmss"
  $SignatureLog_FilePath = Join-Path $PSScriptRoot -ChildPath "$CurrentTime-SignatureDeployment_Details.csv"
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
      $ErrorMessage = $_.Exception.Message
      $DeploymentStatus = "Unsuccessful : $ErrorMessage "
      Write-Host "`n`Signature deployment for $($User.UserPrincipalName) is Unsucessful -Error Occured" -ForegroundColor Red
    }
    $AuditData = [pscustomobject]@{ UserPrincipalName = $User.UserPrincipalName; DeploymentStatus = $DeploymentStatus }
    $AuditData | Export-Csv -Path $SignatureLog_FilePath -NoTypeInformation -Append
    Write-Progress -Activity "Deploying Signature : $($Count) of $($TotalCount) | Current User : $($User.UserPrincipalName)"
    $Count++
  }
  Disconnect_ExchangeOnline_Safely
}

#---------------------------------------------------------------Function to generate and Assign custom text Signature to the users-------------------------------------------------------------------

function Deploy-CustomTextSignature
{
  if (-not $TaskScheduler)
  {
    $Numbers = Get-RequiredFieldsFromUser
  }
  else
  {
    $Numbers = $Inputs[$Index++].Input -split ','
  }
  $UsersCollection = Get-UsersForAssignSignature
  $Count = 1
  $TotalCount = $UsersCollection.count
  $CurrentTime = Get-Date -Format "yyyyMMdd_HHmmss"
  $SignatureLog_FilePath = Join-Path $PSScriptRoot -ChildPath "$CurrentTime-SignatureDeployment_Details.csv"
  $UsersFields = @{ '1' = 'DisplayName'; '2' = 'UserPrincipalName'; '3' = 'MobilePhone'; '4' = 'Phone'; '5' = 'Fax'; '6' = 'Department'; '7' = 'Title'; '8' = 'Office'; '9' = 'Address'; '10' = 'StreetAddress'; '11' = 'City'; '12' = 'StateOrProvince'; '13' = 'PostalCode'; '14' = 'CountryOrRegion' }
  foreach ($User in $UsersCollection)
  {
    try
    {
      $UserDetails = Get-User -Identity $User.UserPrincipalName -ErrorAction Stop
      $UserTextSignature = ""
      foreach ($Number in $Numbers) {
        if ($UsersFields[$Number]) {
          if ($Number -eq "9") {
            $Address = Generate-UserAddress ($UserDetails)
            if ($Address -ne "") {
              $UserTextSignature += "$($Address)<br>"
            }
          }
          elseif ($UserDetails.($UsersFields[$Number]) -ne "") {
            $UserTextSignature += "$($UserDetails.($UsersFields[$Number]))<br>"
          }
        }
      }
      Set-MailboxMessageConfiguration -Identity $UserDetails.UserPrincipalName -SignatureHTML $UserTextSignature -AutoAddSignature $true -AutoAddSignatureOnMobile $true -AutoAddSignatureOnReply $true -ErrorAction Stop
      $DeploymentStatus = "Successful"
    }
    catch
    {
      $ErrorMessage = $_.Exception.Message
      $DeploymentStatus = "Unsuccessful : $ErrorMessage "
      Write-Host "`n`Signature deployment for $($User.UserPrincipalName) is Unsucessful -Error Occured" -ForegroundColor Red
    }
    $AuditData = [pscustomobject]@{ UserPrincipalName = $User.UserPrincipalName; DeploymentStatus = $DeploymentStatus }
    $AuditData | Export-Csv -Path $SignatureLog_FilePath -NoTypeInformation -Append
    Write-Progress -Activity "Deploying Signature : $($Count) of $($TotalCount) | Current User : $($User.UserPrincipalName)"
    $Count++
  }
  Disconnect_ExchangeOnline_Safely

}

#---------------------------------------------------------------Function to generate and Assign custom HTML Signature to the users-------------------------------------------------------------------

function Deploy-CustomHTMLSignature
{
  if (-not $TaskScheduler)
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
          if ($ValueInHTML -ne "" )
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
    $HTMLSignature | Select-Object @{ Name = 'Input'; Expression = { $_ } } | Export-Csv -Path $InputsFolderPath -NoTypeInformation -Append
  }
  else
  {
    $HTMLSignature = $Inputs[$Index++].Input
    if ($HTMLSignature.Contains('%%Address%%'))
    {
      $AddressUsed = $true
    }
  }
  $UsersCollection = Get-UsersForAssignSignature
  $Count = 1
  $TotalCount = $UsersCollection.count
  $CurrentTime = Get-Date -Format "yyyyMMdd_HHmmss"
  $SignatureLog_FilePath = Join-Path $PSScriptRoot -ChildPath "$CurrentTime-SignatureDeployment_Details.csv"
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
      $ErrorMessage = $_.Exception.Message
      $DeploymentStatus = "Unsuccessful : $ErrorMessage "
      Write-Host "`n`Signature deployment for $($User.UserPrincipalName) is Unsucessful -Error Occured" -ForegroundColor Red
    }
    $AuditData = [pscustomobject]@{ UserPrincipalName = $User.UserPrincipalName; DeploymentStatus = $DeploymentStatus }
    $AuditData | Export-Csv -Path $SignatureLog_FilePath -NoTypeInformation -Append
    Write-Progress -Activity "Deploying Signature : $($Count) of $($TotalCount) | Current User : $($User.UserPrincipalName)"
    $Count++
  }
  Disconnect_ExchangeOnline_Safely
}

#-------------------function to verify the status of Signature Assign to user and process the auditlog csv file and disconnect the exchangeonline module-------------------------

function Disconnect_ExchangeOnline_Safely
{
  if (-not ($UserConfirmation -eq 1))
  {
    Write-Host "`n`The signature has been added and can be deployed later by enabling 'PostponeRoamingSignatureUntilLater'."
  }
  Write-Host "`n`Script Execution Completed"
  Disconnect-ExchangeOnline -Confirm:$false
  Write-Host "`n`The Signature Deployment Status Report available in:" -NoNewline -ForegroundColor Yellow
  Write-Host $SignatureLog_FilePath
  if(-not $TaskScheduler)
  {
  $Title = "Confirmation"
  $Question = "Do you want to View the Signature Deployment Status Report?"
  $Choices = "&Yes","&No"
  $Decision = $Host.UI.PromptForChoice($Title,$Question,$Choices,1)
  if ($Decision -eq 0) {
    Start-Process $SignatureLog_FilePath
  }
  }
  exit
}

#---------------------------------------Block to call the function as per the user Choice----------------------------

$Output = ""
if (-not $TaskScheduler)
{
  Write-Host "`n`Signature configuration choices:" -ForegroundColor Yellow
  Write-Host @"        
        `n`1. Create a Text Signature using Default Fields 
        `n`2. Create an HTML Signature using an In-built HTML Template
        `n`3. Create a Text Signature with custom Fields
        `n`4. Create an HTML Signature with a user-defined Template 
"@ -ForegroundColor Cyan
  $UserChoice = Read-Host "`n`Enter Your choice"
}
else
{
  $UserChoice = $Inputs[$Index++].Input
}
while ($true)
{
  if (($UserChoice -eq 1 -or $UserChoice -eq 2 -or $UserChoice -eq 3 -or $UserChoice -eq 4) -and (-not $TaskScheduler))
  {
    $UserChoice | Select-Object @{ Name = 'Input'; Expression = { $_ } } | Export-Csv -Path $InputsFolderPath -NoTypeInformation -Append
  }
  if ($UserChoice -eq 1)
  {
    $Output = "`n`Adding Default Text Signature "
    Deploy-DefaultTextSignature
  }
  elseif ($UserChoice -eq 2)
  {
    $Output = "`n`Adding In-build HTML Signature "
    Deploy-InbuiltHTMLSignature
  }
  elseif ($UserChoice -eq 3)
  {
    $Output = "`n`Adding Custom Text Signature "
    Deploy-CustomTextSignature
  }
  elseif ($UserChoice -eq 4)
  {
    $Output = "`n`Adding Custom HTML Signature "
    Deploy-CustomHTMLSignature
  }
  else
  {
    Write-Host "`n`Enter the correct input" -ForegroundColor Red
    $UserChoice = Read-Host
    continue;
  }
}