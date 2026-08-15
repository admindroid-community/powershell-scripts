<#
    Licensing tasks - Microsoft Graph.

    Replaces the MSOnline-based licensing scripts (Get-MsolAccountSku,
    Set-MsolUserLicense) which no longer function.
#>

# Common SKU part numbers -> readable names. Falls back to the part number when unknown.
$script:SkuFriendlyName = @{
    'O365_BUSINESS_ESSENTIALS'           = 'Microsoft 365 Business Basic'
    'O365_BUSINESS_PREMIUM'              = 'Microsoft 365 Business Standard'
    'SPB'                                = 'Microsoft 365 Business Premium'
    'SPE_E3'                             = 'Microsoft 365 E3'
    'SPE_E5'                             = 'Microsoft 365 E5'
    'SPE_F1'                             = 'Microsoft 365 F3'
    'ENTERPRISEPACK'                     = 'Office 365 E3'
    'ENTERPRISEPREMIUM'                  = 'Office 365 E5'
    'STANDARDPACK'                       = 'Office 365 E1'
    'DESKLESSPACK'                       = 'Office 365 F3'
    'EXCHANGESTANDARD'                   = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'                 = 'Exchange Online (Plan 2)'
    'POWER_BI_PRO'                       = 'Power BI Pro'
    'POWER_BI_STANDARD'                  = 'Power BI (free)'
    'PROJECTPROFESSIONAL'                = 'Project Plan 3'
    'PROJECTPREMIUM'                     = 'Project Plan 5'
    'VISIOCLIENT'                        = 'Visio Plan 2'
    'EMS'                                = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                         = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM'                        = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2'                     = 'Microsoft Entra ID P2'
    'MCOEV'                              = 'Microsoft Teams Phone Standard'
    'MCOMEETADV'                         = 'Microsoft 365 Audio Conferencing'
    'FLOW_FREE'                          = 'Power Automate (free)'
    'TEAMS_EXPLORATORY'                  = 'Teams Exploratory'
    'WINDOWS_STORE'                      = 'Windows Store for Business'
    'RIGHTSMANAGEMENT_ADHOC'             = 'Rights Management Adhoc'
}

function Get-SkuDisplayName {
    param([string]$SkuPartNumber)
    if ($script:SkuFriendlyName.ContainsKey($SkuPartNumber)) {
        return $script:SkuFriendlyName[$SkuPartNumber]
    }
    return $SkuPartNumber
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'licensing.sku-summary'
    Name     = 'Licence Inventory'
    Category = 'Licensing'
    Synopsis = 'Purchased vs assigned vs available units for every SKU.'
    Service  = 'Graph'
    Scopes   = @('Organization.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'LicenseExpiryDateReport / Office365 License Reporting And Management (retired Get-MsolAccountSku)'
    Parameters = @(
        @{ Name = 'OnlyWithSurplus'; Type = 'Bool'; Default = $false
           Help = 'Only SKUs with unassigned units - the ones worth reclaiming.' }
    )
    Execute = {
        param($P)

        $skus = Get-MgSubscribedSku -All -ErrorAction Stop

        foreach ($sku in $skus) {
            $enabled  = [int](Get-SafeProperty $sku 'PrepaidUnits.Enabled' 0)
            $warning  = [int](Get-SafeProperty $sku 'PrepaidUnits.Warning' 0)
            $suspended= [int](Get-SafeProperty $sku 'PrepaidUnits.Suspended' 0)
            $consumed = [int](Get-SafeProperty $sku 'ConsumedUnits' 0)
            $available = $enabled - $consumed

            if ($P.OnlyWithSurplus -and $available -le 0) { continue }

            [pscustomobject]@{
                LicenseName    = Get-SkuDisplayName (Get-SafeProperty $sku 'SkuPartNumber')
                SkuPartNumber  = Get-SafeProperty $sku 'SkuPartNumber'
                TotalUnits     = $enabled
                AssignedUnits  = $consumed
                AvailableUnits = $available
                WarningUnits   = $warning
                SuspendedUnits = $suspended
                UtilisationPct = if ($enabled -gt 0) { [math]::Round(($consumed / $enabled) * 100, 1) } else { 0 }
                AppliesTo      = Get-SafeProperty $sku 'AppliesTo'
                CapabilityStatus = Get-SafeProperty $sku 'CapabilityStatus'
                SkuId          = Get-SafeProperty $sku 'SkuId'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'licensing.unused-licenses'
    Name     = 'Unused Licences'
    Category = 'Licensing'
    Synopsis = 'Licences assigned to inactive, disabled or never-signed-in accounts.'
    Service  = 'Graph'
    Scopes   = @('User.Read.All', 'Organization.Read.All', 'AuditLog.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Find Unused Microsoft 365 Licenses'
    Notes    = 'The direct money-saving report: every row is a licence you are paying for that nobody is using.'
    Parameters = @(
        @{ Name = 'InactiveDays'; Type = 'Int'; Default = 90
           Help = 'Treat accounts with no successful sign-in in this many days as inactive.' }
        @{ Name = 'IncludeDisabled'; Type = 'Bool'; Default = $true
           Help = 'Include sign-in-blocked accounts that still hold licences.' }
        @{ Name = 'IncludeNeverSignedIn'; Type = 'Bool'; Default = $true
           Help = 'Include accounts that have never signed in.' }
    )
    Execute = {
        param($P)

        # SKU id -> readable name
        $skuMap = @{}
        foreach ($s in (Get-MgSubscribedSku -All -ErrorAction Stop)) {
            $skuMap[(Get-SafeProperty $s 'SkuId')] = Get-SkuDisplayName (Get-SafeProperty $s 'SkuPartNumber')
        }

        $props = @('Id','DisplayName','UserPrincipalName','AccountEnabled','SignInActivity',
                   'AssignedLicenses','UserType','Department','CreatedDateTime')

        $users  = Get-MgUser -All -Property $props -ErrorAction Stop
        $cutoff = (Get-Date).AddDays(-1 * [int]$P.InactiveDays)

        foreach ($u in $users) {
            $licenses = @(Get-SafeProperty $u 'AssignedLicenses' @())
            if ($licenses.Count -eq 0) { continue }

            $enabled = Get-SafeProperty $u 'AccountEnabled' $true
            $last = Get-SafeProperty $u 'SignInActivity.LastSuccessfulSignInDateTime'
            if (-not $last) { $last = Get-SafeProperty $u 'SignInActivity.LastSignInDateTime' }

            $reason = $null
            if (-not $enabled -and $P.IncludeDisabled) {
                $reason = 'Sign-in blocked'
            } elseif (-not $last) {
                if (-not $P.IncludeNeverSignedIn) { continue }
                $reason = 'Never signed in'
            } elseif ([datetime]$last -le $cutoff) {
                $reason = "Inactive $([int]((Get-Date) - [datetime]$last).TotalDays) days"
            } else {
                continue
            }

            $names = foreach ($l in $licenses) {
                $id = Get-SafeProperty $l 'SkuId'
                if ($id -and $skuMap.ContainsKey($id)) { $skuMap[$id] } else { $id }
            }

            [pscustomobject]@{
                DisplayName       = Get-SafeProperty $u 'DisplayName'
                UserPrincipalName = Get-SafeProperty $u 'UserPrincipalName'
                Reason            = $reason
                AccountEnabled    = $enabled
                UserType          = Get-SafeProperty $u 'UserType'
                Department        = Get-SafeProperty $u 'Department'
                LastSuccessfulSignIn = $last
                DaysInactive      = Get-DaysSince $last
                LicenseCount      = $licenses.Count
                Licenses          = (@($names) -join '; ')
                UserId            = Get-SafeProperty $u 'Id'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'licensing.user-licenses'
    Name     = 'User Licence Assignments'
    Category = 'Licensing'
    Synopsis = 'Which licences each user holds, and whether direct or group-inherited.'
    Service  = 'Graph'
    Scopes   = @('User.Read.All', 'Organization.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Office 365 User License Report / Find M365 User License Assignment Path / M365 Users Group-based License Assignment Report'
    Notes    = 'The assignment path matters: a direct licence overlapping a group-inherited one is double-paid.'
    Parameters = @(
        @{ Name = 'UserPrincipalName'; Type = 'String'
           Help = 'Optional: a single user. Leave blank for the whole tenant.' }
        @{ Name = 'OnlyOverlapping'; Type = 'Bool'; Default = $false
           Help = 'Only rows where the same SKU is assigned both directly and via a group.' }
    )
    Execute = {
        param($P)

        $skuMap = @{}
        foreach ($s in (Get-MgSubscribedSku -All -ErrorAction Stop)) {
            $skuMap[(Get-SafeProperty $s 'SkuId')] = Get-SkuDisplayName (Get-SafeProperty $s 'SkuPartNumber')
        }

        $props = @('Id','DisplayName','UserPrincipalName','AssignedLicenses','LicenseAssignmentStates','Department')

        $users = if ($P.UserPrincipalName) {
            @(Get-MgUser -UserId $P.UserPrincipalName.Trim() -Property $props -ErrorAction Stop)
        } else {
            Get-MgUser -All -Property $props -ErrorAction Stop
        }

        foreach ($u in $users) {
            $states = @(Get-SafeProperty $u 'LicenseAssignmentStates' @())
            if ($states.Count -eq 0) { continue }

            # Group by SKU so direct + inherited for the same SKU can be compared.
            $bySku = $states | Group-Object -Property { Get-SafeProperty $_ 'SkuId' }

            foreach ($grp in $bySku) {
                $skuId = $grp.Name
                $direct   = @($grp.Group | Where-Object { -not (Get-SafeProperty $_ 'AssignedByGroup') })
                $viaGroup = @($grp.Group | Where-Object { Get-SafeProperty $_ 'AssignedByGroup' })

                $isOverlap = ($direct.Count -gt 0 -and $viaGroup.Count -gt 0)
                if ($P.OnlyOverlapping -and -not $isOverlap) { continue }

                $path = if ($isOverlap) { 'Direct + Group (overlap)' }
                        elseif ($viaGroup.Count -gt 0) { 'Group' }
                        else { 'Direct' }

                $groupIds = @($viaGroup | ForEach-Object { Get-SafeProperty $_ 'AssignedByGroup' } | Where-Object { $_ })
                $errors   = @($grp.Group | ForEach-Object { Get-SafeProperty $_ 'Error' } | Where-Object { $_ -and $_ -ne 'None' })

                [pscustomobject]@{
                    DisplayName       = Get-SafeProperty $u 'DisplayName'
                    UserPrincipalName = Get-SafeProperty $u 'UserPrincipalName'
                    Department        = Get-SafeProperty $u 'Department'
                    LicenseName       = if ($skuMap.ContainsKey($skuId)) { $skuMap[$skuId] } else { $skuId }
                    AssignmentPath    = $path
                    IsOverlapping     = $isOverlap
                    AssignedByGroups  = ($groupIds -join '; ')
                    AssignmentErrors  = ($errors -join '; ')
                    SkuId             = $skuId
                    UserId            = Get-SafeProperty $u 'Id'
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'licensing.unlicensed-users'
    Name     = 'Unlicensed Users'
    Category = 'Licensing'
    Synopsis = 'Enabled member accounts holding no licence.'
    Service  = 'Graph'
    Scopes   = @('User.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'List Unlicensed Users in M365'
    Parameters = @(
        @{ Name = 'IncludeGuests'; Type = 'Bool'; Default = $false
           Help = 'Guests are normally unlicensed by design.' }
        @{ Name = 'IncludeDisabled'; Type = 'Bool'; Default = $false
           Help = 'Include sign-in-blocked accounts.' }
    )
    Execute = {
        param($P)

        $props = @('Id','DisplayName','UserPrincipalName','AccountEnabled','AssignedLicenses',
                   'UserType','Department','JobTitle','CreatedDateTime')

        foreach ($u in (Get-MgUser -All -Property $props -ErrorAction Stop)) {
            if (@(Get-SafeProperty $u 'AssignedLicenses' @()).Count -gt 0) { continue }

            $type = Get-SafeProperty $u 'UserType'
            if (-not $P.IncludeGuests -and $type -eq 'Guest') { continue }

            $enabled = Get-SafeProperty $u 'AccountEnabled' $true
            if (-not $P.IncludeDisabled -and -not $enabled) { continue }

            [pscustomobject]@{
                DisplayName       = Get-SafeProperty $u 'DisplayName'
                UserPrincipalName = Get-SafeProperty $u 'UserPrincipalName'
                UserType          = $type
                AccountEnabled    = $enabled
                Department        = Get-SafeProperty $u 'Department'
                JobTitle          = Get-SafeProperty $u 'JobTitle'
                CreatedDateTime   = Get-SafeProperty $u 'CreatedDateTime'
                UserId            = Get-SafeProperty $u 'Id'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'licensing.set-user-license'
    Name     = 'Assign / Remove Licence'
    Category = 'Licensing'
    Synopsis = 'Adds or removes a SKU for a user.'
    Service  = 'Graph'
    Scopes   = @('User.ReadWrite.All', 'Organization.Read.All')
    Risk     = 'Write'
    Replaces = 'Manage Microsoft 365 Licenses using MS Graph (retired Set-MsolUserLicense)'
    Parameters = @(
        @{ Name = 'UserPrincipalName'; Type = 'String'; Required = $true
           Help = 'Target user.' }
        @{ Name = 'SkuPartNumber'; Type = 'String'; Required = $true
           Help = 'SKU part number, e.g. SPE_E3. See licensing.sku-summary.' }
        @{ Name = 'Operation'; Type = 'Choice'; Default = 'Add'; Options = @('Add','Remove')
           Help = 'Assign or remove the licence.' }
        @{ Name = 'WhatIf'; Type = 'Bool'; Default = $true
           Help = 'Report the intended change without applying it.' }
    )
    Execute = {
        param($P)

        $upn = $P.UserPrincipalName.Trim()
        $sku = (Get-MgSubscribedSku -All -ErrorAction Stop |
                Where-Object { (Get-SafeProperty $_ 'SkuPartNumber') -eq $P.SkuPartNumber })

        if (-not $sku) {
            throw "SKU '$($P.SkuPartNumber)' is not present in this tenant. Run licensing.sku-summary to list valid SKUs."
        }

        $skuId = Get-SafeProperty $sku 'SkuId'
        $available = [int](Get-SafeProperty $sku 'PrepaidUnits.Enabled' 0) - [int](Get-SafeProperty $sku 'ConsumedUnits' 0)

        if ($P.Operation -eq 'Add' -and $available -le 0) {
            throw "No available units for '$($P.SkuPartNumber)' ($available free)."
        }

        if ($P.WhatIf) {
            return [pscustomobject]@{
                UserPrincipalName = $upn
                Operation         = $P.Operation
                LicenseName       = Get-SkuDisplayName $P.SkuPartNumber
                SkuPartNumber     = $P.SkuPartNumber
                AvailableUnits    = $available
                Status            = 'WhatIf - no change made'
            }
        }

        $addParam    = if ($P.Operation -eq 'Add')    { @(@{ SkuId = $skuId }) } else { @() }
        $removeParam = if ($P.Operation -eq 'Remove') { @($skuId) }             else { @() }

        try {
            Set-MgUserLicense -UserId $upn -AddLicenses $addParam -RemoveLicenses $removeParam -ErrorAction Stop | Out-Null
            $status = 'Success'
        } catch {
            $status = "Failed: $($_.Exception.Message)"
        }

        [pscustomobject]@{
            UserPrincipalName = $upn
            Operation         = $P.Operation
            LicenseName       = Get-SkuDisplayName $P.SkuPartNumber
            SkuPartNumber     = $P.SkuPartNumber
            AvailableUnits    = $available
            Status            = $status
        }
    }
}
