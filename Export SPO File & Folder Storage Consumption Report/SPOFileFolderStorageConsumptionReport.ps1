<#
=============================================================================================
Name:           Export SharePoint Online File & Folder Storage Usage Report Using PowerShell 
Version:        2.0
Website:        o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Retrieves file/folder consumption details for all document libraries in a site. 
2. Exports file/folder consumption details for a list of sites. 
3. Exports the ‘SPO file storage consumption report’ and ‘SPO folder storage consumption report’ into a CSV file. 
4. File/folder size can be exported in preferred units such as MB, KB, B, and GB. 
5. Automatically installs the PnP PowerShell module (if not installed already) upon your confirmation. 
6. The script can be executed with an MFA-enabled account too. 
7. The script is scheduler friendly. 
8. The script uses modern authentication to connect SharePoint Online. 
9. It can be executed with certificate-based authentication (CBA) too. 

For detailed script execution: https://o365reports.com/2024/08/07/export-sharepoint-online-file-folder-storage-usage-report-using-powershell/

Change Log:

V1.0 (Aug 08, 2024)- Script created
V2.0 (Jan 15, 2025)- Due to PnP app regsitration removed, added support for passing ClientId during script execution.
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$AdminName ,
    [string]$Password, 
    [String]$ClientId,
    [String]$CertificateThumbprint,
    [String]$TenantName,
    [string]$SiteUrl,
    [string]$ImportCsv,
    [ValidateSet('GB', 'MB', 'KB', 'B')]
    [string]$Unit
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
    elseif($TenantName -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
    {
        Connect-PnPOnline -Url $Url -ClientId $ClientId -Thumbprint $CertificateThumbprint  -Tenant "$TenantName.onmicrosoft.com" 
    }
    else
    {
        Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
    }
}
Function Convert-Bytes {
    $Bytes = switch ($Unit.ToUpper()) {
        'GB' { 1GB }
        'MB' { 1MB }
        'KB' { 1KB }
        'B'  { 1 }
        Default {
            Write-Host "Invalid unit is provided. csv will display file sizes in megabyte"
            1MB
        }
    }
    return $Bytes
}
Function HumanReadableByteSize ($Size) {
    Switch ($Size) {
	{$_ -gt 1TB} {$Size = ($Size / 1TB).ToString("n2") + " TB";break}
	{$_ -gt 1GB} {$Size = ($Size / 1GB).ToString("n2") + " GB";break}
	{$_ -gt 1MB} {$Size = ($Size / 1MB).ToString("n2") + " MB";break}
	{$_ -gt 1KB} {$Size = ($Size / 1KB).ToString("n2") + " KB";break}
	default {$Size =  "$Size B"}
	}
Return $Size
}
Function Calculate-Percentages($OverAllSiteConsumption){
     $ListConsumption = 0
     $CsvData = Import-Csv -Path $ResultCsv_Folder
     Foreach($line in $CsvData){
       If($line.'Folder Name' -eq "ROOT"){
            $ListConsumption = $line.'OverallSize'
        }
        $PercentageByParentLibrary = If($ListConsumption -eq 0){0} Else{ "{0:N4}" -f  (($line.'OverallSize'/$ListConsumption)*100)}
        $PercentageByOverAllSite = If($OverAllSiteConsumption -eq 0){0} Else{ "{0:N4}" -f  (($line.'OverallSize'/$OverAllSiteConsumption)*100)}
        $line | Add-Member -MemberType NoteProperty -Name "Percentage By Parent Library" -Value $PercentageByParentLibrary -Force
        $line | Add-Member -MemberType NoteProperty -Name "Percentage By OverAll Site" -Value $PercentageByOverAllSite -Force
        $line.PSObject.Properties.Remove("OverallSize")
     }
     $CsvData| Export-Csv -Path $ResultCsv_Folder -NoTypeInformation
}

Function Get-FileStorageConsumption($Item){
     $size = $Item.FieldValues.SMTotalFileStreamSize/$bytes
     $FileItemInfo = [PSCustomObject]@{
            "Site"         = $Site
            "Library Name" = $List.Title
            "URL" = $Item.FieldValues.FileRef
            "File Name" = $Item.FieldValues.FileLeafRef
            "File Size" = HumanReadableByteSize ($size*$bytes)
            "File Size ($Unit)" = "{0:N2}" -f ($Size)
            "Created By" = $Item.FieldValues.Author.Email
            "Created Date" = $Item.FieldValues.Created
            "Last Modified By" = $Item.FieldValues.Editor.Email
            "Last Modified Date" = $Item.FieldValues.Modified  
            }
     $FileItemInfo | Export-csv -Path $ReportOutput_File -Append -NoTypeInformation -Force
}

Function Get-FolderStorageConsumption($Object){
    $OverallSize = 0
    $FoldersSize = 0
    $FilesSize = 0

    $Items = Get-PnPListItem -List $List -PageSize 2000 | where {$_.FieldValues.FileDirRef -eq $Object.FieldValues.FileRef}

    Foreach($Item in $Items){
        Write-Progress -Activity ("Site : "+$Site + "  | List : "+$List.Title) -Status ("Processing Item: "+$Item.FieldValues.FileLeafRef)
        If($Item.FileSystemObjectType -eq "Folder"){
            $FoldersSize += Get-FolderStorageConsumption $Item
        }
        Else{
            $FilesSize += $Item.FieldValues.SMTotalFileStreamSize
            Get-FileStorageConsumption $Item
        }
    }
    $OverallSize = $FoldersSize + $FilesSize
    $ItemInfo = [PSCustomObject]@{
                    "Site"         = $Site
                    "Library Name" = $List.Title
                    "Folder Name" = $Object.FieldValues.FileLeafRef
                    "Location" = $Object.FieldValues.FileRef
                    "Total Size" = (HumanReadableByteSize ($OverallSize))
                    "Direct Folders Size" = (HumanReadableByteSize ($FoldersSize))
                    "Direct Files Size" = (HumanReadableByteSize ($FilesSize))
                    "Total Size($unit)" = "{0:N2}" -f($OverallSize/$bytes)
                    "Direct Folders Size($Unit)" = "{0:N2}" -f($FoldersSize/$bytes)
                    "Direct Files Size($Unit)" = "{0:N2}" -f($FilesSize/$bytes)
                    "OverallSize" = $OverallSize
                    
                } 
    $ItemInfo | Export-CSV $HelperCsv  -NoTypeInformation -Append
    return $OverallSize
}

Function Get-SPOStorageConsumption{

    $ExcludedLists = @("Form Templates","Style Library","Site Assets","Site Pages", "Preservation Hold Library", "Pages", "Images",
                                "Site Collection Documents", "Site Collection Images")
    $DocumentLibraries = Get-PnPList | Where-Object {$_.Hidden -eq $False -and $_.Title -notin $ExcludedLists -and $_.BaseType -eq "DocumentLibrary"}
    $OverAllSiteConsumption = 0
    Foreach($List in $DocumentLibraries){
        $Objects = Get-PnPListItem -List $List -PageSize 2000 | where {$_.FieldValues.FileDirRef -eq $List.RootFolder.ServerRelativeUrl}

        $OverallSize = 0
        $FoldersSize = 0
        $FilesSize = 0

        Foreach($Object in $Objects){
            Write-Progress -Activity ("Site : "+$Site + "  | List : "+$List.Title) -Status ("Processing Item: "+$Object.FieldValues.FileLeafRef)
            If($Object.FileSystemObjectType -eq "Folder"){
                $FoldersSize +=  Get-FolderStorageConsumption $Object    
            }
            Else{
                $FilesSize += $Object.FieldValues.File_x0020_Size
                Get-FileStorageConsumption $Object
            }
        }
        $OverallSize = $FoldersSize + $FilesSize
        $OverAllSiteConsumption += $OverallSize
        $ItemInfo = [PSCustomObject]@{
                    "Site"         = $Site
                    "Library Name" = $List.Title
                    "Folder Name" = "ROOT"
                    "Location" = $List.RootFolder.ServerRelativeUrl
                    "Total Size" = (HumanReadableByteSize ($OverallSize))
                    "Direct Folders Size" = (HumanReadableByteSize ($FoldersSize))
                    "Direct Files Size" = (HumanReadableByteSize ($FilesSize))
                    "Total Size($unit)" = "{0:N2}" -f($OverallSize/$bytes)
                    "Direct Folders Size($Unit)" = "{0:N2}" -f($FoldersSize/$bytes)
                    "Direct Files Size($Unit)" = "{0:N2}" -f($FilesSize/$bytes)
                    "OverallSize" = $OverallSize
                } 

        $ItemInfo | Export-CSV $ResultCsv_Folder  -NoTypeInformation -Append -Force
         if (Test-Path $HelperCsv -PathType Leaf) {
            $HelperData = Import-Csv -Path $HelperCsv
            $HelperData | Export-Csv $ResultCsv_Folder -NoTypeInformation -Append -force
            Remove-Item $HelperCsv
         }
    }
    if((Test-Path -Path $ResultCsv_Folder) -eq "True") 
    {
        Calculate-Percentages $OverAllSiteConsumption
    }

}

Installation-Module
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$HelperCsv = "$PsScriptRoot\Helper$timestamp.csv"
$ResultCsv_Folder = "$PsScriptRoot\FolderStorageConsumption$timestamp.csv"
$ReportOutput_Folder = "$PsScriptRoot\SPO-Folder Storage Consumption Report $timestamp.csv"
$ReportOutput_File = "$PsScriptRoot\SPO-File Storage Consumption Report $timestamp.csv"

 If($Unit -eq ""){
    Write-Host -f cyan "Note: By default file sizes will be shown in MB."
    Write-Host -f cyan "You can specify the unit[B, KB, MB, GB] by passing the desired unit using '-Unit' argument."
    $Unit = "MB"
    $Bytes = 1MB
 }Else{
    $Bytes = Convert-Bytes
 }

If($ImportCsv -ne ""){
    $ListOfSites = Import-csv -Path $ImportCsv 
    Foreach($Site in $ListOfSites){
        Connection-Module -Url $Site.SiteUrl
        $OverAllSiteConsumption = 0
        $Site = (Get-PnPWeb | Select Title).Title
        Get-SPOStorageConsumption
        if (Test-Path $ResultCsv_Folder -PathType Leaf) {
            $ResultCsv_FolderData = Import-Csv -Path $ResultCsv_Folder
            $ResultCsv_FolderData| Export-Csv $ReportOutput_Folder -NoTypeInformation -Append -force
            Remove-Item $ResultCsv_Folder
        }
        Disconnect-PnPOnline
    }

}
Else{
    If($SiteUrl -eq ""){
        $SiteUrl = Read-Host "Site Url "
    }
    Connection-Module -Url $SiteUrl
    $Site = (Get-PnPWeb | Select Title).Title 
    Get-SPOStorageConsumption
    if (Test-Path $ResultCsv_Folder -PathType Leaf) {
            $ResultCsv_FolderData = Import-Csv -Path $ResultCsv_Folder
            $ResultCsv_FolderData| Export-Csv $ReportOutput_Folder -NoTypeInformation -Append -force
            Remove-Item $ResultCsv_Folder
    }
    Disconnect-PnPOnline
}

if((Test-Path -Path $ReportOutput_Folder ) -eq "True") {
    write-host `n Two files have been created -NoNewline 
    write-host " SPO-Folder Storage Consumption $timestamp.csv " -NoNewline -f Yellow
    write-host and -NoNewline 
    write-host " SPO-Folder Storage Consumption $timestamp.csv" -f yellow
    Write-Host `n The Output files are availble in:   -NoNewline -ForegroundColor Yellow
    Write-Host $PsScriptRoot 
    Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
    Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
    $Prompt = New-Object -ComObject wscript.shell   
    $UserInput = $Prompt.popup("Do you want to open output files?",`   
     0,"Open Output File",4)   
    If ($UserInput -eq 6)   
    {   
        Invoke-Item $ReportOutput_File
        Invoke-Item $ReportOutput_Folder              
   } 
}
 