<#
=============================================================================================
Name:         Export SharePoint Online File Version History Report Using PowerShell    
Version:      1.0
website:      o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Retrieves file version history for all documents in a site. 
2. Exports file version history for a list of sites. 
3. Exports version history information for files uploaded by a specific user. 
4. Finds files with larger version history based on your input. 
5. File size can be exported in preferred units such as MB, KB, B, and GB. 
6. Automatically installs the PnP PowerShell module (if not installed already) upon your confirmation. 
7. The script can be executed with an MFA-enabled account too. 
8. Exports report results as a CSV file. 
9. The script is scheduler friendly. 
10. It can be executed with certificate-based authentication (CBA) too. 

For detailed Script execution: https://o365reports.com/2024/06/20/export-sharepoint-online-file-version-history-report-using-powershell/
============================================================================================
#>
Param
(
    [Parameter(Mandatory = $false)]
    [string]$AdminName,
    [string]$Password,
    [String] $ClientId,
    [String] $CertificateThumbprint,
    [String] $TenantName,
    [string]$SiteUrl,
    [int]$VersionCount = -1,
    [string]$ImportCsv,
    [string]$UserId,
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
        Connect-PnPOnline -Url $Url -interactive
           
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
Function Store-FileVersionInformation($FileItem){
     write-host $fileItem.FieldValues.FileRef -f Green
    If($Versions.Count -ge $VersionCount){
        $VersionSize = $Versions | Measure-Object -Property Size -Sum | Select-Object -expand Sum
        $VersionSize = [Math]::Round(($VersionSize/$bytes),2)
        $FileSize = [Math]::Round(($FileItem.FieldValues.File_x0020_Size/$Bytes),2)
        $TotalFileSize = $FileSize + $VersionSize

        $FileVersionData =  [PSCustomObject][Ordered]@{
            "Site" = $Site
            "Library" = $List.Title
            "File Name"  = $FileItem.FieldValues.FileLeafRef
            "File URL" = $File.ServerRelativeUrl
            "Major Versions" = $File.MajorVersion
            "Minor Versions" = $file.MinorVersion
            "Versions Count" =  $Versions.Count
            "File Size" = HumanReadableByteSize ($FileSize*$Bytes)
            "Version History Size" = HumanReadableByteSize ($VersionSize*$Bytes)
            "Total File Size" = HumanReadableByteSize ($TotalFileSize*$Bytes)
            "File Size ($unit)" = $FileSize
            "Version History Size ($unit) " = $VersionSize
            "Total File Size ($unit)" = $TotalFileSize
            "Created By" = $FileItem.FieldValues.Author.Email
            "Created Date" = $File.TimeCreated
            "Modified By" = $FileItem.FieldValues.Editor.Email
            "Last Modified" = $file.TimeLastModified
            
        }
        $FileVersionData | Export-Csv -Path $ReportOutput -NoTypeInformation -Append -force
        $Global:ItemCount++
    }        
}
Function Get-SiteFileVersion($DocumentLibraries){
    ForEach($List in $DocumentLibraries)
    {
        $Files = Get-PnPListItem -List $List -PageSize 2000 -Fields File_x0020_Size, FileRef, Author | Where {$_.FileSystemObjectType -eq "File"}
        Foreach($FileItem in $Files){
            Write-Progress -Activity ("Site : "+$Site +" | List : "+$List.Title) -Status ("Processing Item: "+$FileItem.FieldValues.FileLeafRef)
            $File = Get-PnPProperty -ClientObject $FileItem  -Property File
            $Versions = Get-PnPProperty -ClientObject $File -Property Versions
            If($UserId -ne "" -and $UserId -eq $FileItem.FieldValues.Author.Email){
                Store-FileVersionInformation $FileItem
            }
            ElseIf($UserId -eq ""){
                Store-FileVersionInformation $FileItem
            }
        }
    }
}

If($Unit -ne ""){
    $Bytes = Convert-Bytes
}
else{
    Write-Host -f cyan "Note: By default file sizes will be shown in MB."
    Write-Host -f cyan "You can specify the unit[B, KB, MB, GB] by passing the desired unit using '-Unit' argument."
    $Bytes = 1MB
    $Unit = "MB"
}
Installation-Module

$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ReportOutput = "$PSScriptRoot\SPO File Version History Report $timestamp.csv"
$Global:ItemCount = 0

If($ImportCsv -ne ""){
    $ListOfSites = Import-csv -Path $ImportCsv 
    Foreach($Site in $ListOfSites){
        Connection-Module -Url $Site.SiteUrl
        $Site = (Get-PnPWeb | Select Title).Title
        $ExcludedLists = @("Form Templates","Style Library","Site Assets","Site Pages", "Preservation Hold Library", "Pages", "Images",
                   "Site Collection Documents", "Site Collection Images")
        $DocumentLibraries = Get-PnPList | Where-Object {$_.Hidden -eq $False -and $_.Title -notin $ExcludedLists -and $_.BaseType -eq "DocumentLibrary"}
        Get-SiteFileVersion $DocumentLibraries
        Disconnect-PnPOnline
    }

}
Else{
    If($SiteUrl -eq ""){
        $SiteUrl = Read-Host "Site Url " 
    }

    Connection-Module -Url $SiteUrl
    $Site = (Get-PnPWeb | Select Title).Title 
    $ExcludedLists = @("Form Templates","Style Library","Site Assets","Site Pages", "Preservation Hold Library", "Pages", "Images",
                       "Site Collection Documents", "Site Collection Images")
    $DocumentLibraries = Get-PnPList | Where-Object {$_.Hidden -eq $False -and $_.Title -notin $ExcludedLists -and $_.BaseType -eq "DocumentLibrary"}
    Get-SiteFileVersion $DocumentLibraries
    Disconnect-PnPOnline
}


if((Test-Path -Path $ReportOutput) -eq "True") 
{
    Write-Host `nThe output file contains $Global:ItemCount files
    Write-Host `n The Output file availble in: -NoNewline -ForegroundColor Yellow
    Write-Host $OutputCSV 
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


