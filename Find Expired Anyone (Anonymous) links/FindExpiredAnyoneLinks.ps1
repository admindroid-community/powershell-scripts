<#
=============================================================================================
Name:           Get All Expired Anyone Links in SharePoint Online Using PowerShell  
Version:        1.0
Website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Exports all expired anyone links in your SPO environment. 
2. Exports expired anyone links for a list of sites. 
3. Automatically installs the PnP PowerShell module (if not installed already) upon your confirmation. 
4. The script can be executed with an MFA-enabled account too. 
5. Exports report results as a CSV file. 
6. The script is scheduler friendly. 
7. The script uses modern authentication to connect SharePoint Online. 
8. It can be executed with certificate-based authentication (CBA) too. 

For detailed script execution: https://o365reports.com/2024/07/16/get-all-expired-anyone-links-in-sharepoint-online-using-powershell/

============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$AdminName ,
    [string]$Password ,
    [String]$ClientId ,
    [String]$CertificateThumbprint,
    [string]$TenantName,
    [string]$ImportCsv 
)
Function Installation-Module{
    $Module = Get-InstalledModule -Name PnP.PowerShell -RequiredVersion 1.12.0 -ErrorAction SilentlyContinue
    If($Module -eq $null){
        Write-Host PnP PowerShell Module is not available -ForegroundColor Yellow
        $Confirm = Read-Host Are you sure you want to install module? [Yy] Yes [Nn] No
        If($Confirm -match "[yY]") { 
            Write-Host "Installing PnP PowerShell module..."
            Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Force -AllowClobber -Scope CurrentUser
            Import-Module -Name Pnp.Powershell -RequiredVersion 1.12.0           
        } 
        Else{ 
           Write-Host PnP PowerShell module is required to connect SharePoint Online.Please install module using Install-Module PnP.PowerShell cmdlet. 
           Exit
        }
    }
    Write-Host `nConnecting to SharePoint Online...
} 
Function Connection-Module{
    param
    (
        [Parameter(Mandatory = $true)]
        [String] $Url
    )
    if(($AdminName -ne "") -and ($Password -ne ""))
    {
        $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $Credential  = New-Object System.Management.Automation.PSCredential $AdminName,$SecuredPassword
        Connect-PnPOnline -Url $Url -Credential $Credential
    }
    elseif($ClientId -ne "" -and $CertificateThumbprint -ne "" -and $TenantName -ne "")
    {
        Connect-PnPOnline -Url $Url -ClientId $ClientId -Thumbprint $CertificateThumbprint  -Tenant "$TenantName.onmicrosoft.com" 
    }
    else
    {
        Connect-PnPOnline -Url $Url -Interactive   
    }
}
Function Get-SharedLinkInfo($ListItems) {
    $Ctx = Get-PnPContext
    ForEach ($Item in $ListItems) {
    Write-Progress -Activity ("Site Name: $Site") -Status ("Processing Item: "+ $Item.FieldValues.FileLeafRef)
        $HasUniquePermissions = Get-PnPProperty -ClientObject $Item -Property HasUniqueRoleAssignments

        If ($HasUniquePermissions) {        
            $SharingInfo = [Microsoft.SharePoint.Client.ObjectSharingInformation]::GetObjectSharingInformation($Ctx, $Item, $false, $false, $false, $true, $true, $true, $true)
            $Ctx.Load($SharingInfo)
            $Ctx.ExecuteQuery()

            ForEach ($ShareLink in $SharingInfo.SharingLinks) {           
                If ($ShareLink.Url -and $ShareLink.LinkKind -like "*Anonymous*") { 
                    $LinkStatus = $false                        
                    $LinkCreated = ([DateTime]$ShareLink.Created).tolocalTime()

                    $CurrentDateTime = Get-Date
                    If($ShareLink.Expiration -ne ""){
                        $Expiration = ([DateTime]$ShareLink.Expiration).tolocalTime()
                        If($Expiration -lt $CurrentDateTime){
                            $daysExpired = ($currentDateTime - $expiration).Days
                            $LinkStatus = $true
                        } 
                    }
                    If($LinkStatus){
                        If($ShareLink.IsEditLink)
                        {
                            $AccessType="Write"
                        }
                        ElseIf($shareLink.IsReviewLink)
                        {
                            $AccessType="Review"
                        }
                        Else
                        {
                            $AccessType="Read"
                        }
                        $Results = [PSCustomObject]@{
                            "Site Name"             = $Site
                            "Library"          = $List.Title
                            "File Name"             = $Item.FieldValues.FileLeafRef
                            "File URL"         = $Item.FieldValues.FileRef
                            "Access Type"      = $AccessType
                            "File Type"         = $Item.FieldValues.File_x0020_Type 
                            "Link Expired Date"  = $Expiration
                            "Days Since Expired"     = $daysExpired
                            "Link Created Date "    = $LinkCreated   
                            "Last Modified On "   = ([DateTime]$ShareLink.LastModified).tolocalTime()                          
                            "Shared Link"       = $ShareLink.Url  
                        }
                        $Results | Export-CSV  -path $ReportOutput -NoTypeInformation -Append  -Force
                        $Global:ItemCount++
                    }                    
                }
            }
        }
    }
}
Function Get-SharedLinks{
    $ExcludedLists = @("Form Templates","Style Library","Site Assets","Site Pages", "Preservation Hold Library", "Pages", "Images",
                       "Site Collection Documents", "Site Collection Images")
    $DocumentLibraries = Get-PnPList | Where-Object {$_.Hidden -eq $False -and $_.Title -notin $ExcludedLists -and $_.BaseType -eq "DocumentLibrary"}
    Foreach($List in $DocumentLibraries){
        $ListItems = Get-PnPListItem -List $List -PageSize 2000  | Where {$_.FileSystemObjectType -eq "File"}
        Get-SharedLinkInfo $ListItems
    }
}
Installation-Module
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ReportOutput = "$PSScriptRoot\Expired_SPO_Shared_Links $timestamp.csv"
$Global:ItemCount = 0
If($TenantName -eq ""){
    $TenantName = Read-Host "Enter your tenant name (e.g., 'contoso' for 'contoso.onmicrosoft.com')"
}

If($ImportCsv -ne ""){
    $SiteCollections = Import-Csv -Path $ImportCsv
    Foreach($Site in $SiteCollections){
        Connection-Module -Url $Site.SiteUrl
        $Site = (Get-PnPWeb | Select Title).Title
        Get-SharedLinks
        Disconnect-PnPOnline 
    }
}
Else{
    Connection-Module -Url "https://$TenantName-admin.sharepoint.com"
    $SiteCollections = Get-PnPTenantSite  | Where -Property Template -NotIn ("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")
    Disconnect-PnPOnline -WarningAction SilentlyContinue
    ForEach($Site in $SiteCollections)
    {
        Connection-Module -Url $Site.Url  
        $Site = (Get-PnPWeb | Select Title).Title
        Get-SharedLinks
        Disconnect-PnPOnline -WarningAction SilentlyContinue 
    }
}
if((Test-Path -Path $ReportOutput) -eq "True") 
{
    Write-Host `nThe output file contains $Global:ItemCount files
    Write-Host `n The Output file availble in:  -NoNewline -ForegroundColor Yellow
    Write-Host $ReportOutput
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    $Prompt = New-Object -ComObject wscript.shell   
    $UserInput = $Prompt.popup("Do you want to open output file?",`   
    0,"Open Output File",4)   
    If ($UserInput -eq 6)   
    {   
        Invoke-Item "$ReportOutput"   
    } 
}
else{
    Write-Host -f Yellow "No Records Found"
}