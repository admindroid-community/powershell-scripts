<#
=============================================================================================
Name:           Connect to MS Graph PowerShell using Certificate
Description:    This script automates Azure app registration
For detailed Script execution: https://blog.admindroid.com/connect-to-microsoft-graph-powershell-using-certificate/
============================================================================================
#>
param (
    $TenantID =$null,
    $ClientID = $null,
    $CertificateThumbprint = $null
) 
Function ConnectMgGraphModule
{
    $MsGraphModule =  Get-Module Microsoft.Graph -ListAvailable
    if($MsGraphModule -eq $null)
    { 
        Write-host "Important: Microsoft graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install Microsoft graph module? [Y] Yes [N] No  
        if($confirm -match "[yY]") 
        { 
            Write-host "Installing Microsoft graph module..."
            Install-Module Microsoft.Graph -Scope CurrentUser
            Write-host "Microsoft graph module is installed in this machine successfully" -ForegroundColor Magenta 
        } 
        else
        { 
            Write-host "Exiting. `nNote: Microsoft graph module must be available in your system to run the script" -ForegroundColor Red
            Exit 
        } 
    }
    Connect-MgGraph -Scopes "Application.ReadWrite.All,Directory.ReadWrite.All"  -ErrorAction SilentlyContinue -Errorvariable ConnectionError |Out-Null
    if($ConnectionError -ne $null)
    {
        Write-Host "$ConnectionError" -Foregroundcolor Red
        Exit
    }
    Write-Host "Microsoft Graph Powershell module is connected successfully" -ForegroundColor Green
    $Script:TenantID = (Get-MgOrganization).Id
}
function RegisterApplication
{
    Write-Progress -Activity "Registering an application"
    while(1)
    {
        $Script:AppName = Read-Host "`nEnter a name for the new App"
        if($AppName -eq "")
        {
            Write-Host "You didn't enter any name. Please provide an app name to continue." -ForegroundColor Red
            continue
        }
        break
    }
    $Script:RedirectURI = "https://login.microsoftonline.com/common/oauth2/nativeclient"
    $params = @{
        DisplayName = $AppName
        SignInAudience="AzureADMyOrg"
        PublicClient=@{
                RedirectUris = "$RedirectURI"
        }
        RequiredResourceAccess = @(
            @{
                ResourceAppId = "00000003-0000-0000-c000-000000000000" #  Microsoft Graph Resource ID
                ResourceAccess = @(
                    @{
                        Id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61" #Directory.Read.All -> Id
                        Type = "Role"                               #Role -> Application permission
                     }
                    )
                }
            )
    }
    try{
        $Script:App = New-MgApplication -BodyParameter $params 
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        CloseConnection
    }
    Write-Host "`nApp created successfully" -ForegroundColor Green
    $Script:APPObjectID = $App.Id
    $Script:APPID = $App.AppId
}
function CertificateCreation
{
    Write-Progress -Activity "Creating certificate"
    $Script:CertificateName = "$AppName-Keycertifcate"   
    $Script:path = "Cert:\CurrentUser\My\"
    $Script:CertificatePath = Get-ChildItem -Path $path
    $Script:Subject = "CN=$CertificateName"
    try
    {
        $Script:Certificate = New-SelfSignedCertificate -Subject $Subject -CertStoreLocation "$path" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -HashAlgorithm SHA256 
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        CloseConnection
    }
    Write-Host "`nCertificate created successfully" -ForegroundColor Green
}
Function ImportCertificate
{
    Write-Progress -Activity "Importing certificate"
    $UploadCertificate = Read-Host "Enter the certificate path(For example: C:\Users\Admin\user.cer)"
    try
    {
        $Script:Certificate = Import-Certificate -FilePath "$UploadCertificate" -CertStoreLocation "Cert:\CurrentUser\My" 
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        $ImportCertificateError = $True
        ShowAppDetails
        CloseConnection
    }
    Write-Host "`nCertificate imported successfully" -ForegroundColor Green
}
Function UpdateCertficateDetails
{
    $GetAppInfo= Get-MgApplication -ApplicationId $APPObjectID -Property KeyCredentials
    $Thumbprint=$Certificate.Thumbprint
    $CertificateName=$Certificate.Subject
    if($GetAppInfo -ne $null)
    {
        $NewKeys+=$GetAppInfo.KeyCredentials 
        UploadCertificate
        # Get new certificate key credentials after updating new certificate
        Start-Sleep -Milliseconds 10000 # Takes time to update certificate to an application
        while(1)
        {
            $GetAppInfo= Get-MgApplication -ApplicationId $APPObjectID -Property KeyCredentials
            if($GetAppInfo.KeyCredentials -eq $null)
            {
                Write-Host "The certificate is not yet uploaded to an application. Waiting for 3 seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds 3
                Continue
            }
            break
        }
        $NewKeys+=$GetAppInfo.KeyCredentials  #new keycredential
        Update-MgApplication -ApplicationId $APPObjectID -KeyCredentials $NewKeys
    }
    else 
    {
        Write-Host "The application does not exist. Please try again with a different name." -ForegroundColor Green
    }
}
function UploadCertificate 
{
    Write-Progress -Activity "Uploading certificate"
    $KeyCredential = @{
        Type  = "AsymmetricX509Cert";
        Usage = "Verify";
        key   = $Certificate.RawData
    }
    Update-MgApplication -ApplicationId $APPObjectID  -KeyCredentials $KeyCredential -ErrorAction SilentlyContinue -ErrorVariable ApplicationError
    if($ApplicationError -ne $null)
    {
        Write-Host "$ApplicationError" -ForegroundColor Red
        CloseConnection 
    }
    Write-Host "`nCertificate uploaded successfully" -ForegroundColor Green
    $Script:Thumbprint = $Certificate.Thumbprint
}
function SecureCertificate
{
    Write-Progress -Activity "Exporting pfx certificate" 
    $Script:CertificateLocation = Get-Location
    $Script:CertificateLocation = "$CertificateLocation\$CertificateName.pfx"
    $GetPassword = Read-Host "`nPlease enter password to secure your certificate"
    try
    {
        $script:ExportError="False"
        $MyPwd = ConvertTo-SecureString -String "$GetPassword" -Force -AsPlainText
        Export-PfxCertificate -Cert "Cert:\CurrentUser\My\$Thumbprint" -FilePath "$CertificateLocation" -Password $MyPwd |Out-Null
    }
    catch
    {
        $script:ExportError="True"
        Write-Host $_.Exception.Message  -ForegroundColor Red
        return
    }
    Write-Host "`nPfx file exported successfully" -ForegroundColor Green
}
function GrantPermission
{
    Write-Progress -Activity "Granting admin consent..."
    Start-Sleep -Seconds 20
    $Script:ClientID = $App.AppId
    $URL = "https://login.microsoftonline.com/$TenantID/adminconsent"
    $Url="$URL`?client_id=$ClientID"
    Write-Host "`nMS Graph requires admin consent to access data. Please grant access to the application" -ForegroundColor Cyan
    While(1)
    {
        Add-Type -AssemblyName System.Windows.Forms
        $script:mainForm = New-Object System.Windows.Forms.Form -Property @{
            Width  = 680
            Height = 640
        }
        $script:webBrowser = New-Object System.Windows.Forms.WebBrowser -Property @{
            Width  = 680
            Height = 640
            URL    = $URL
        }
        $document={
            if($webBrowser.Url -eq "$RedirectURI`?admin_consent=True&tenant=$TenantID" -or $webBrowser.Url -match "error")
            {
                $mainForm.Close()
            }
            if($webBrowser.DocumentText.Contains("We received a bad request"))
            {
                $mainForm.Close()
            }
        }
        $webBrowser.ScriptErrorsSuppressed = $true
        $webBrowser.Add_DocumentCompleted($document)
        $mainForm.Controls.Add($webBrowser)
        $mainForm.Add_Shown({ $mainForm.Activate() ;$mainForm.Refresh()})
        [void] $mainForm.ShowDialog()
        if($webBrowser.Url.AbsoluteUri -eq "$RedirectURI`?admin_consent=True&tenant=$TenantID")
        {
            Write-Host "`nAdmin consent granted successfully" -ForegroundColor Green
            break
        }
        else
        {
            Write-Host "`nAdmin consent failed." -ForegroundColor Red
            $Confirm = Read-Host "Do you want to retry admin consent? [Y] Yes [N] No"
            if($confirm -match "[yY]")
            {
                Continue
            } 
            else
            {
                Write-Host "You can grant admin consent manually in azure portal." -ForegroundColor Yellow
                break
            }
        }
    }
}
Function ShowAppDetails
{
    Write-Host "`nApp info:" -ForegroundColor Magenta
    $GetAppInfo = Get-MgApplication|?{$_.AppId -eq "$APPID"}
    $Owner = Get-MgApplicationOwner -ApplicationId $GetAppInfo.Id|Select-Object -ExpandProperty AdditionalProperties
    $Script:CertificateList = $GetAppInfo.KeyCredentials
    $AppInfo=[pscustomobject]@{'App Name'              = $GetAppInfo.DisplayName
                               'Application(Client) Id'        = $GetAppInfo.AppId
                               'Object Id' = $GetAppInfo.Id
                               'Tenant Id'             = $TenantID
                               'CertificateThumbprint' = $Thumbprint
                               'App Created Date Time' = $GetAppInfo.CreatedDateTime
                               'App Owner'             = (@($Owner.displayName)| Out-String).Trim()
    }
    if($ImportCertificateError -eq $True)
    {
        $AppInfo | select 'App Name','Application(Client) Id','Object Id','Tenant Id','App Created Date Time','App Owner'|fl
        Write-Host "You can copy & save" -NoNewline
        Write-Host " Client Id" -ForegroundColor Cyan -NoNewline
        Write-Host " and try again to add certificate to your application."
        Return
    }
    if($Action -ne 5)
    {
        $AppInfo | select 'App Name','Application(Client) Id','Object Id','Tenant Id','CertificateThumbprint','App Created Date Time','App Owner'|fl
        Write-Host "You can copy and save" -NoNewline
        Write-Host " Client Id, Tenant Id, Certificate ThumbPrint" -ForegroundColor Cyan -NoNewline
        Write-Host " as they will be required to connect to MS Graph PowerShell using certificate-based authentication."
    }
    else
    {
        $AppInfo | select 'App Name','Application(Client) Id','Object Id','Tenant Id','App Created Date Time','App Owner'|fl
        Write-Host "App certificates  :"
        $CertificateList | Format-Table -Property DisplayName, KeyId, StartdateTime, EndDateTime

    }
}
Function RevokeCertificate
{
    Write-Progress -Activity "Revoking certificate"
    $NewKeys = @()
    $APPID = Read-Host "Enter an application ID(Client ID) you want to revoke the certificate for that app"
    $GetAppInfo = Get-MgApplication |?{$_.AppId -eq "$APPID"}
    if($GetAppInfo -ne $null)
    {
        ShowAppDetails
        $KeyId=Read-Host "`nEnter the certificate key ID to revoke that certificate"
        if($GetAppInfo.KeyCredentials.KeyId -notcontains($KeyId))
        {
            Write-Host "Certificate not found" -ForegroundColor Red
            CloseConnection
        }
        foreach($List in $CertificateList)
        {
            if($List.KeyId -eq "$KeyId")
            {
                continue
            }
            $NewKeys+=$List
        }
        Update-MgApplication -ApplicationId $GetAppInfo.Id -KeyCredentials $NewKeys 
        Write-Host "`nCertificate revoked successfully" -ForegroundColor Green
    }
    else 
    {
        Write-Host "The application does not exist." -ForegroundColor Red
        CloseConnection
    }
}
Function ConnectApplication
{
    Write-Progress -Activity "Connecting MS Graph"
    try
    {
        if($ParameterPassed -eq "False")
        {
            $TenantID = Read-Host "`nPlease provide the tenant ID of the application"
            $ClientID = Read-Host "Please provide the client ID of the application"
            While($true){
                $CertificatePath = Read-Host "Please provide certificate path(.cer or .pfx)"
                $TestPath = Test-Path -Path $CertificatePath -PathType Leaf
                if($TestPath -eq $false)
                {
                    Write-Host "The certificate file could not be found." -ForegroundColor Red
                    Continue
                }
            }
            $CheckExtension = (Get-ChildItem "$CertificatePath").Extension
            if($CheckExtension -eq ".pfx")
            { 
                $Password = Read-Host "Please enter password to import certificate"
                $MyPwd = ConvertTo-SecureString -String "$Password" -Force -AsPlainText
                $LoadCertificate = Import-PfxCertificate -FilePath $CertificatePath -CertStoreLocation "Cert:\CurrentUser\My" -Password $MyPwd
            }
            else
            {
                $LoadCertificate = Import-Certificate -FilePath "$CertificatePath" -CertStoreLocation "Cert:\CurrentUser\My"
            }
            $CertificateThumbprint = $LoadCertificate.Thumbprint
        }
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Exit
    }
    Connect-MgGraph -TenantId $TenantID -ClientId $ClientID -CertificateThumbprint $CertificateThumbprint -ErrorAction SilentlyContinue -ErrorVariable ApplicationConnectionError
    if($ApplicationConnectionError -ne $null)
    {
        Write-Host $ApplicationConnectionError -ForegroundColor Red
        Exit
    }
    
    Get-MgContext
}
function CloseConnection
{
   Disconnect-MgGraph|Out-Null 
   Exit
}
$ParameterPassed="False"
if($TenantID -ne $null -and $ClientID -ne $null -and $CertificateThumbprint -ne $null)
{
    $ParameterPassed="True"
    ConnectApplication
    Exit
}
Write-Host "`nWe can perform below operations." -ForegroundColor Cyan
Write-Host "           1. Register an app with new certificate" -ForegroundColor Yellow -NoNewline
Write-Host " (Creates application - >Adds new certificate -> Grants admin consent -> Exports certificate)" -ForegroundColor White
Write-Host "           2. Register an app with existing certificate" -ForegroundColor Yellow -NoNewline
Write-Host " (Creates application -> Adds existing certificate -> Grants admin consent)" -ForegroundColor White
Write-Host "           3. Add certificate to an existing application" -ForegroundColor Yellow 
Write-Host "           4. Connect MgGraph" -ForegroundColor Yellow
Write-Host "           5. Revoke certificate" -ForegroundColor Yellow
$Action=Read-Host "`nPlease choose the action to continue" 
switch($Action){
   1 {  
   Write-Host "`nConnecting to MS Graph to create a application..."
        ConnectMgGraphModule
        RegisterApplication
        CertificateCreation
        UploadCertificate
        SecureCertificate
        GrantPermission
        ShowAppDetails
        if($ExportError -ne "True")
        {
            Write-Host "`nYour Pfx certificate is available in $CertificateLocation" -ForegroundColor Green
        }
        break

   }
   2 {
        Write-Host "`nConnecting to MS Graph to create a application.."
        ConnectMgGraphModule
        RegisterApplication
        ImportCertificate
        UploadCertificate
        GrantPermission
        ShowAppDetails
        break
   }
   3 {
        ConnectMgGraphModule
        $APPID = Read-Host "Enter the application id(Client id) of the app:"
        $AppInfo = Get-MgApplication |?{$_.AppId -eq "$APPID"}
        if($AppInfo -eq $null)
        {
            Write-Host "Application not found." -ForegroundColor Red
            CloseConnection
        }
        $AppName = $AppInfo.DisplayName
        $APPObjectID =$AppInfo.Id
        $confirm= Read-Host Are you sure you want to create new certificate? [Y] Yes [N] No. Select '"No"' to import existing certificate.  
        if($confirm -match "[yY]") 
        { 
            CertificateCreation
        } 
        else
        { 
            ImportCertificate
        } 
        UpdateCertficateDetails
        ShowAppDetails
        break
   }
   4 {
        ConnectApplication
        Exit
   }
   5 {
        ConnectMgGraphModule
        RevokeCertificate
        break
   }
   Default {
        Write-Host "No Action Found" -ForegroundColor Red
        Exit
    }
}
CloseConnection