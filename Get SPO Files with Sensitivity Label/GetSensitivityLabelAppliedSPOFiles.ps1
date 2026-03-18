<#
=================================================================================
Name: Get All Files with Sensitivity Labels in SharePoint Online
Version: 1.0
Website: o365reports.com
~~~~~~~~~~~~~~~~~
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. Exports all SharePoint Online files with sensitivity labels across your environment. 
2. Automatically checks for the PnP PowerShell module and installs it with your confirmation if it is missing. 
3. The script can also register an Microsoft Entra ID app for PnP authentication with the required permissions, if needed. 
4. Supports CSV input to analyze specific sites only. 
5. Retrieves files with a specific sensitivity label. 
6. Includes filters for label assignment type such as default, manual, auto assigned, and unknown. 
7. The script can be executed with an MFA enabled account too. 
8. The script is scheduler-friendly and compatible with certificate-based authentication (CBA). 

For detailed script execution: https://o365reports.com/find-files-with-sensitivity-labels-in-sharepoint-online-using-powershell/
=================================================================================
#>
Param(
    [Parameter(Mandatory = $false)]
    [string]$AdminName,
    [string]$Password,
    [string]$CertificateThumbprint,
    [string]$ClientId,
    [string]$SharePointDomainName,
	[string]$AppName,
    [switch]$RegisterAppForPnP,
    [string]$ImportCsv,
	[string]$Url,
	[string]$LabelName,
    [switch]$DefaultLabel,
	[switch]$ManuallyAssignedLabel,       
    [switch]$AutoAssignedLabel,
	[switch]$UnknownLabel
)

 function Installation-Module {
    try {
        $Module = Get-InstalledModule -Name PnP.PowerShell -MinimumVersion 2.12.0 -ErrorAction SilentlyContinue
        if ($null -eq $Module) {
            Write-Host "SharePoint PnP PowerShell Module is not available" -ForegroundColor Yellow
            $Confirm = Read-Host "Are you sure you want to install module? [Y]es [N]o"
            if ($Confirm -match "[yY]") { 
                Write-Host "Installing PnP PowerShell module..."
                Install-Module PnP.PowerShell -Force -AllowClobber -Scope CurrentUser        
            } 
            else {
                Write-Host "PnP PowerShell module is required to connect SharePoint Online. Please install module using 'Install-Module PnP.PowerShell' cmdlet."
                Exit
            }
            Import-Module -Name PnP.PowerShell
        }
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Exit
    }
       
}

#Connection-Module
Function Connection-Module {
    param([string]$Url)
    if ($AdminName -and $Password -and $ClientId) {
        $SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential ($AdminName, $SecurePassword)
        Connect-PnPOnline -Url $Url -Credential $Credential -ClientId $ClientId
    } elseif ($SharePointDomainName -and $ClientId -and $CertificateThumbprint -and $Url) {
        Connect-PnPOnline -Url $Url -ClientId $ClientId -Thumbprint $CertificateThumbprint -Tenant "$SharePointDomainName.onmicrosoft.com"
    } else {
        Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
    }
}

Function RegisterAppInEntra
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$SharePointDomainName
    )
    if($AppName -eq "") {
        $AppName = Read-Host "Enter application name"
    }
	try{
		$RegisteredApp = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName $AppName -Tenant "$SharePointDomainName.onmicrosoft.com" -GraphDelegatePermissions Sites.FullControl.All,InformationProtectionPolicy.Read
		return $RegisteredApp.'AzureAppId/ClientId'
	}
	catch{
		Write-Host "`n$($_.Exception.Message)" -ForegroundColor Red
		Exit
	}
	
}

Function Convert-FileSize {
    param([long]$bytes)

    if ($bytes -ge 1TB) {
        return ("{0:N2} TB" -f ($bytes / 1TB))
    }
    elseif ($bytes -ge 1GB) {
        return ("{0:N2} GB" -f ($bytes / 1GB))
    }
    elseif ($bytes -ge 1MB) {
        return ("{0:N2} MB" -f ($bytes / 1MB))
    }
    elseif ($bytes -ge 1KB) {
        return ("{0:N2} KB" -f ($bytes / 1KB))
    }
    else {
        return ("{0} Bytes" -f $bytes)
    }
}



Function Export-LabeledItemsReport {
	
	$SupportedExtensions = @(
        "doc","docx","docm","dot","dotx","dotm","xls","xlsx","xlt","xla","xlc","xlm","xlw",
        "xltx","xlsm","xltm","xlam","xlsb","ppt","pptx","pot","pps","ppa","ppsx","ppsxm","potx","ppam",
        "pptm","potm","ppsm","pdf"
    )

    # Excluded library list
    $ExcludedLibraries = @(
        "Form Templates", "Preservation Hold Library", "Site Assets", "Site Pages", "Images", "Pages", 
        "Settings", "Videos", "Site Collection Documents", "Site Collection Images", "Style Library", 
        "AppPages", "Apps for SharePoint", "Apps for Office"
    )

    Get-PnPList -ErrorAction SilentlyContinue | Where-Object { $_.Hidden -eq $false -and $_.Title -notin $ExcludedLibraries -and $_.BaseType -eq "DocumentLibrary" } | ForEach-Object {
	    $List = $_
        Get-PnPListItem -List $List.Title -PageSize 2000 | Where-Object { $_.FileSystemObjectType -eq "File"} |
        ForEach-Object {
			$FieldValues = $_.FieldValues
            $LabelId = $FieldValues["_IpLabelId"]
            $HasLabel = -not [string]::IsNullOrWhiteSpace($LabelId)
            if (($FieldValues["FSObjType"] -ne 0) -or -not $HasLabel) { return }
			$FileType = $FieldValues["File_x0020_Type"]
			if ($SupportedExtensions -contains $FileType) {
				$Label = $FieldValues["_DisplayName"]
				$FileUrl = $FieldValues["FileRef"]
				if ($FieldValues["FSObjType"] -eq 0 -and -not [string]::IsNullOrEmpty($FileUrl)) {
					try {
						$SensitivityLabelInfo = Get-PnPFileSensitivityLabel -Url $FileUrl
						$LabelAssignmentMethod = $SensitivityLabelInfo.AssignmentMethod
							
					}
					catch {
						Write-Host $_.Exception.Message
					}			
				}
			}
			$ShouldExport = $false
		

		# Mode-based filtering
		if ($ManuallyAssignedLabel -and -not [string]::IsNullOrEmpty($LabelAssignmentMethod)) {
			$ShouldExport = ($LabelAssignmentMethod -eq "privileged")
		}
		elseif ($AutoAssignedLabel -and -not [string]::IsNullOrEmpty($LabelAssignmentMethod)) {
			$ShouldExport = ($LabelAssignmentMethod -eq "auto")
		}
		elseif ($DefaultLabel -and -not [string]::IsNullOrEmpty($LabelAssignmentMethod)) {
			$ShouldExport = ($LabelAssignmentMethod -eq "standard")
		}
		elseif ($UnknownLabel -and -not [string]::IsNullOrEmpty($LabelAssignmentMethod)){
			$ShouldExport = ($LabelAssignmentMethod -eq "unknownFutureValue")
		}
		elseif ($LabelName -ne "") {
			if ($Label -and ($Label -eq $LabelName) -and ($null -ne $LabelId)) {
				$ShouldExport = $true
			}
		}
		else {
			$ShouldExport = -not [string]::IsNullOrWhiteSpace($Label)
		}
		# Export if eligible
		if ($ShouldExport) {
			$Result = [PSCustomObject]@{
				"Site Name"             = $Script:Site
				"Site URL"				= $Script:SiteUrl
				"Library"               = $List.Title
				"File Name"             = $FieldValues["FileLeafRef"]
                "File URL"              = $FieldValues["FileRef"]
                "Sensitivity Label Name" = $Label
				"Label Assignment Method"= $LabelAssignmentMethod
				"Label ID"              = $LabelId
				"File Type"             = $FieldValues["File_x0020_Type"]
				"File Size"             = Convert-FileSize $FieldValues["File_x0020_Size"]
				"Modified By"           = ($FieldValues["Modified_x0020_By"] -split "\|")[-1]
				"Created By"            = ($FieldValues["Created_x0020_By"] -split "\|")[-1]
			}

			$Result | Export-Csv -Path $Global:ReportOutput -Append -NoTypeInformation
			$Global:ItemCount++
		}
		}
	}
}

# Start Execution
Installation-Module
$TimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$Global:ReportOutput = "$(Get-Location)\SPO_File_Sensitivity_Label_Report_$TimeStamp.csv"
$Global:ItemCount = 0

if($RegisterAppForPnP.IsPresent)
{
    if($SharePointDomainName -eq "")
    {
        $SharePointDomainName = Read-Host "`nEnter your SharePoint tenant name (e.g., 'contoso' for 'contoso.sharepoint.com')"
    }    
    $ClientId = RegisterAppInEntra -TenantName $SharePointDomainName
}

if($Url -eq "" -and $ImportCsv -eq "" -and $SharePointDomainName -eq "")
{
    $SharePointDomainName = Read-Host "Enter your SharePoint tenant name (e.g., 'contoso' for 'contoso.sharepoint.com')"
}
if($ClientId -eq ""){
	$ClientId = Read-host "Enter your ClientId" 
}

# Process Sites
if ($Url) {
    if( $Url -notlike "*-my.sharepoint.com*") {
        Connection-Module -Url $Url
        $Script:Site = (Get-PnPWeb -ErrorAction SilentlyContinue).Title
	    Export-LabeledItemsReport
    }
    else {
        Write-Host "`nOneDrive for Business site '$($Url)' is not supported. Please provide a SharePoint Online site URL to continue." -ForegroundColor Yellow
    }
}
elseif ($ImportCsv) {
    Import-Csv -Path $ImportCsv |
    ForEach-Object {
		Write-Progress -Activity "Processing $($_.SitesUrl)" 
        $Script:SiteUrl = $_.SiteUrl
        if( $Script:SiteUrl -notlike "*-my.sharepoint.com*") {
            Connection-Module -Url $Script:SiteUrl
            $Script:Site = (Get-PnPWeb -ErrorAction SilentlyContinue).Title
		    Export-LabeledItemsReport
        }
        else {
            Write-Host "`nOneDrive for Business site '$($Script:SiteUrl)' is not supported." -ForegroundColor Yellow
        }
    }
    Disconnect-PnPOnline -WarningAction SilentlyContinue
}
else {
    Connection-Module -Url "https://$SharePointDomainName-admin.sharepoint.com"

    Get-PnPTenantSite |
    Where-Object { $_.Template -notin @("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1") } |
    ForEach-Object {
		Write-Progress -Activity "Processing $($_.Url)" 
		$Script:SiteUrl = $_.Url
        Connection-Module -Url $_.Url
        $Script:Site = (Get-PnPWeb -ErrorAction SilentlyContinue).Title
		Export-LabeledItemsReport
    }
    Disconnect-PnPOnline -WarningAction SilentlyContinue
}
 
# Report Summary
if (Test-Path -Path $ReportOutput) {
    Write-Host `nThe output file contains $Global:ItemCount entries.
    Write-Host "`n The Output file availble in: " -NoNewline -ForegroundColor Yellow
    Write-Host $ReportOutput`n
    $Prompt = New-Object -ComObject wscript.shell   
    $UserInput = $Prompt.popup("Do you want to open output file?",0,"Open Output File",4)   
    If ($UserInput -eq 6)   
    {   
        Invoke-Item "$ReportOutput"   
    } 
} else {
    Write-Host "`nNo Records Found"
}

 Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
 Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green