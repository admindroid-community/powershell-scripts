<#
    SharePoint / OneDrive sharing governance and usage reporting.
#>

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'sharing.onedrive-usage'
    Name     = 'OneDrive Usage'
    Category = 'Sharing'
    Synopsis = 'Per-user OneDrive storage consumption and activity.'
    Service  = 'Graph'
    Scopes   = @('Reports.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'OneDrive Usage Report / Export OneDrive Urls and Usage Size'
    Notes    = 'Uses the Graph usage reports endpoint, which returns CSV; it is parsed back into objects here.'
    Parameters = @(
        @{ Name = 'Period'; Type = 'Choice'; Default = 'D30'
           Options = @('D7','D30','D90','D180'); Help = 'Reporting window.' }
        @{ Name = 'MinimumSizeGB'; Type = 'Int'; Default = 0
           Help = 'Only accounts above this consumption.' }
    )
    Execute = {
        param($P)

        $tmp = [System.IO.Path]::GetTempFileName()
        try {
            Get-MgReportOneDriveUsageAccountDetail -Period $P.Period -OutFile $tmp -ErrorAction Stop
            $rows = Import-Csv -Path $tmp -ErrorAction Stop
        } finally {
            Remove-Item $tmp -ErrorAction SilentlyContinue
        }

        foreach ($r in $rows) {
            $usedBytes = 0
            $usedRaw = Get-SafeProperty $r 'Storage Used (Byte)'
            if ($usedRaw) { [void][int64]::TryParse($usedRaw, [ref]$usedBytes) }
            $usedGB = ConvertTo-FriendlySize $usedBytes

            if ([int]$P.MinimumSizeGB -gt 0 -and ($null -eq $usedGB -or $usedGB -lt [int]$P.MinimumSizeGB)) { continue }

            $allocBytes = 0
            $allocRaw = Get-SafeProperty $r 'Storage Allocated (Byte)'
            if ($allocRaw) { [void][int64]::TryParse($allocRaw, [ref]$allocBytes) }

            [pscustomobject]@{
                OwnerDisplayName  = Get-SafeProperty $r 'Owner Display Name'
                OwnerPrincipalName= Get-SafeProperty $r 'Owner Principal Name'
                SiteUrl           = Get-SafeProperty $r 'Site URL'
                IsDeleted         = Get-SafeProperty $r 'Is Deleted'
                LastActivityDate  = Get-SafeProperty $r 'Last Activity Date'
                FileCount         = Get-SafeProperty $r 'File Count'
                ActiveFileCount   = Get-SafeProperty $r 'Active File Count'
                StorageUsedGB     = $usedGB
                StorageAllocatedGB= ConvertTo-FriendlySize $allocBytes
                UsagePercent      = if ($allocBytes -gt 0) { [math]::Round(($usedBytes / $allocBytes) * 100, 1) } else { $null }
                ReportPeriodDays  = Get-SafeProperty $r 'Report Period'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'sharing.sharepoint-usage'
    Name     = 'SharePoint Site Usage'
    Category = 'Sharing'
    Synopsis = 'Site-level storage, file counts and last activity.'
    Service  = 'Graph'
    Scopes   = @('Reports.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Export SPO File & Folder Storage Consumption Report / SPO Document Library Report'
    Parameters = @(
        @{ Name = 'Period'; Type = 'Choice'; Default = 'D30'
           Options = @('D7','D30','D90','D180'); Help = 'Reporting window.' }
        @{ Name = 'InactiveDays'; Type = 'Int'; Default = 0
           Help = 'Only sites with no activity for this many days.' }
    )
    Execute = {
        param($P)

        $tmp = [System.IO.Path]::GetTempFileName()
        try {
            Get-MgReportSharePointSiteUsageDetail -Period $P.Period -OutFile $tmp -ErrorAction Stop
            $rows = Import-Csv -Path $tmp -ErrorAction Stop
        } finally {
            Remove-Item $tmp -ErrorAction SilentlyContinue
        }

        foreach ($r in $rows) {
            $lastActivity = Get-SafeProperty $r 'Last Activity Date'
            $idle = Get-DaysSince $lastActivity

            if ([int]$P.InactiveDays -gt 0) {
                if ($null -ne $idle -and $idle -lt [int]$P.InactiveDays) { continue }
            }

            $usedBytes = 0
            $usedRaw = Get-SafeProperty $r 'Storage Used (Byte)'
            if ($usedRaw) { [void][int64]::TryParse($usedRaw, [ref]$usedBytes) }

            [pscustomobject]@{
                SiteUrl          = Get-SafeProperty $r 'Site URL'
                OwnerDisplayName = Get-SafeProperty $r 'Owner Display Name'
                SiteName         = Get-SafeProperty $r 'Site Name'
                IsDeleted        = Get-SafeProperty $r 'Is Deleted'
                LastActivityDate = $lastActivity
                DaysInactive     = $idle
                FileCount        = Get-SafeProperty $r 'File Count'
                ActiveFileCount  = Get-SafeProperty $r 'Active File Count'
                PageViewCount    = Get-SafeProperty $r 'Page View Count'
                VisitedPageCount = Get-SafeProperty $r 'Visited Page Count'
                StorageUsedGB    = ConvertTo-FriendlySize $usedBytes
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'sharing.anonymous-links'
    Name     = 'Anonymous Sharing Links'
    Category = 'Sharing'
    Synopsis = 'Anyone-with-the-link shares across a site, with expiry state.'
    Service  = 'PnP'
    Risk     = 'ReadOnly'
    Replaces = 'Get All Anonymously Shared Links / Find Expired Anyone (Anonymous) links / Get All Sharing Links in SharePoint Online'
    Notes    = 'Connect with -SharePointUrl pointing at the site to scan. Anonymous links are the highest-exposure sharing type.'
    Parameters = @(
        @{ Name = 'ListTitle'; Type = 'String'; Default = 'Documents'
           Help = 'Document library to scan.' }
        @{ Name = 'OnlyExpired'; Type = 'Bool'; Default = $false
           Help = 'Only links whose expiry date has passed.' }
    )
    Execute = {
        param($P)

        $items = Get-PnPListItem -List $P.ListTitle -PageSize 500 -ErrorAction Stop

        foreach ($item in $items) {
            $fileRef = Get-SafeProperty $item 'FieldValues.FileRef'
            if (-not $fileRef) { continue }

            $sharing = $null
            try {
                $sharing = Get-PnPFileSharingLink -Identity $fileRef -ErrorAction Stop
            } catch {
                continue   # not shared, or not a file
            }

            foreach ($link in @($sharing)) {
                $scope = Get-SafeProperty $link 'Link.Scope'
                if ($scope -ne 'anonymous') { continue }

                $expiry = Get-SafeProperty $link 'ExpirationDateTime'
                $isExpired = $false
                if ($expiry) {
                    try { $isExpired = ([datetime]$expiry -lt (Get-Date)) } catch { }
                }
                if ($P.OnlyExpired -and -not $isExpired) { continue }

                [pscustomobject]@{
                    FilePath       = $fileRef
                    FileName       = Get-SafeProperty $item 'FieldValues.FileLeafRef'
                    LinkScope      = $scope
                    LinkType       = Get-SafeProperty $link 'Link.Type'
                    WebUrl         = Get-SafeProperty $link 'Link.WebUrl'
                    ExpirationDate = $expiry
                    IsExpired      = $isExpired
                    HasPassword    = Get-SafeProperty $link 'HasPassword'
                    ShareId        = Get-SafeProperty $link 'Id'
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'sharing.external-users'
    Name     = 'SharePoint External Users'
    Category = 'Sharing'
    Synopsis = 'Guests with access to SharePoint and OneDrive content.'
    Service  = 'Graph'
    Scopes   = @('User.Read.All', 'Sites.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Export SharePoint Online External Users report'
    Execute = {
        param($P)

        $guests = Get-MgUser -All -Filter "userType eq 'Guest'" `
                    -Property 'Id','DisplayName','UserPrincipalName','Mail','CreatedDateTime','ExternalUserState' `
                    -ErrorAction Stop

        foreach ($g in $guests) {
            $upn  = Get-SafeProperty $g 'UserPrincipalName'
            $mail = Get-SafeProperty $g 'Mail'

            # The guest UPN encodes the source tenant domain before #EXT#.
            $sourceDomain = $null
            if ($upn -and $upn -match '^(.+?)_(.+?)#EXT#') { $sourceDomain = $Matches[2] }
            elseif ($mail -and $mail -match '@(.+)$')      { $sourceDomain = $Matches[1] }

            [pscustomobject]@{
                DisplayName       = Get-SafeProperty $g 'DisplayName'
                UserPrincipalName = $upn
                Mail              = $mail
                SourceDomain      = $sourceDomain
                InvitationState   = Get-SafeProperty $g 'ExternalUserState'
                CreatedDateTime   = Get-SafeProperty $g 'CreatedDateTime'
                AgeDays           = Get-DaysSince (Get-SafeProperty $g 'CreatedDateTime')
                UserId            = Get-SafeProperty $g 'Id'
            }
        }
    }
}
