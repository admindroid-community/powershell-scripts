<#
    Identity tasks - Microsoft Graph.

    Several of these replace scripts in the original collection that depend on the
    retired MSOnline / AzureAD modules.
#>

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'identity.mfa-status'
    Name     = 'MFA Registration Status'
    Category = 'Identity'
    Synopsis = 'Per-user MFA registration state, capabilities and default method.'
    Service  = 'Graph'
    Scopes   = @('AuditLog.Read.All', 'UserAuthenticationMethod.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Office 365 User MFA Status Report (retired MSOnline StrongAuthenticationMethods)'
    Notes    = 'Uses the bulk authentication-method registration report: one paged call for the whole tenant rather than a Graph round trip per user.'
    Parameters = @(
        @{ Name = 'OnlyUnregistered'; Type = 'Bool'; Default = $false
           Help = 'Return only users who have not registered any MFA method.' }
        @{ Name = 'OnlyAdmins'; Type = 'Bool'; Default = $false
           Help = 'Return only users holding a privileged directory role.' }
    )
    Execute = {
        param($P)

        $details = Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop

        if ($P.OnlyAdmins) {
            $details = $details | Where-Object { Get-SafeProperty $_ 'IsAdmin' $false }
        }
        if ($P.OnlyUnregistered) {
            $details = $details | Where-Object { -not (Get-SafeProperty $_ 'IsMfaRegistered' $false) }
        }

        foreach ($d in $details) {
            $methods = @(Get-SafeProperty $d 'MethodsRegistered' @())

            [pscustomobject]@{
                DisplayName        = Get-SafeProperty $d 'UserDisplayName'
                UserPrincipalName  = Get-SafeProperty $d 'UserPrincipalName'
                IsAdmin            = Get-SafeProperty $d 'IsAdmin' $false
                MfaRegistered      = Get-SafeProperty $d 'IsMfaRegistered' $false
                MfaCapable         = Get-SafeProperty $d 'IsMfaCapable' $false
                SsprRegistered     = Get-SafeProperty $d 'IsSsprRegistered' $false
                SsprCapable        = Get-SafeProperty $d 'IsSsprCapable' $false
                PasswordlessCapable= Get-SafeProperty $d 'IsPasswordlessCapable' $false
                DefaultMethod      = Get-SafeProperty $d 'DefaultMfaMethod'
                MethodsRegistered  = ($methods -join '; ')
                MethodCount        = $methods.Count
                LastUpdated        = Get-SafeProperty $d 'LastUpdatedDateTime'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'identity.mfa-methods-detail'
    Name     = 'MFA Methods (per user detail)'
    Category = 'Identity'
    Synopsis = 'Every registered authentication method for one or more named users.'
    Service  = 'Graph'
    Scopes   = @('UserAuthenticationMethod.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'List M365 Users Registered MFA Authentication Methods'
    Parameters = @(
        @{ Name = 'UserPrincipalName'; Type = 'String'; Required = $true
           Help = 'One or more UPNs, comma separated.' }
    )
    Execute = {
        param($P)

        $upns = @($P.UserPrincipalName -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

        foreach ($upn in $upns) {
            try {
                $methods = Get-MgUserAuthenticationMethod -UserId $upn -ErrorAction Stop
            } catch {
                Write-M365Log -Level Warning -Source 'identity.mfa-methods-detail' -Message "$upn : $($_.Exception.Message)"
                continue
            }

            foreach ($m in $methods) {
                # The method type is carried in the OData type discriminator.
                $odata = Get-SafeProperty $m 'AdditionalProperties' @{}
                $type  = if ($odata -is [System.Collections.IDictionary] -and $odata.Contains('@odata.type')) {
                             ($odata['@odata.type'] -replace '#microsoft.graph.', '')
                         } else { 'unknown' }

                [pscustomobject]@{
                    UserPrincipalName = $upn
                    MethodType        = $type
                    MethodId          = Get-SafeProperty $m 'Id'
                    Detail            = if ($odata -is [System.Collections.IDictionary]) {
                                            (($odata.GetEnumerator() |
                                              Where-Object { $_.Key -notlike '@odata*' } |
                                              ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; ')
                                        } else { $null }
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'identity.inactive-users'
    Name     = 'Inactive Users'
    Category = 'Identity'
    Synopsis = 'Licensed accounts with no successful sign-in in N days.'
    Service  = 'Graph'
    Scopes   = @('User.Read.All', 'AuditLog.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Get M365 Inactive User Report / Last Successful Sign-in Date Report'
    Notes    = 'signInActivity requires Entra ID P1/P2. Graph cannot combine a signInActivity filter with other filters, so filtering is done client side.'
    Parameters = @(
        @{ Name = 'InactiveDays'; Type = 'Int'; Default = 90
           Help = 'Days since last successful sign-in.' }
        @{ Name = 'IncludeNeverSignedIn'; Type = 'Bool'; Default = $true
           Help = 'Include accounts with no recorded sign-in at all.' }
        @{ Name = 'LicensedOnly'; Type = 'Bool'; Default = $true
           Help = 'Only accounts holding a licence (these are the ones costing money).' }
    )
    Execute = {
        param($P)

        $props = @('Id','DisplayName','UserPrincipalName','AccountEnabled','SignInActivity',
                   'AssignedLicenses','UserType','CreatedDateTime','Department','JobTitle')

        $users = Get-MgUser -All -Property $props -ErrorAction Stop
        $cutoff = (Get-Date).AddDays(-1 * [int]$P.InactiveDays)

        foreach ($u in $users) {
            $licenses = @(Get-SafeProperty $u 'AssignedLicenses' @())
            if ($P.LicensedOnly -and $licenses.Count -eq 0) { continue }

            $lastSuccess = Get-SafeProperty $u 'SignInActivity.LastSuccessfulSignInDateTime'
            if (-not $lastSuccess) { $lastSuccess = Get-SafeProperty $u 'SignInActivity.LastSignInDateTime' }

            if (-not $lastSuccess) {
                if (-not $P.IncludeNeverSignedIn) { continue }
            } else {
                if ([datetime]$lastSuccess -gt $cutoff) { continue }
            }

            [pscustomobject]@{
                DisplayName       = Get-SafeProperty $u 'DisplayName'
                UserPrincipalName = Get-SafeProperty $u 'UserPrincipalName'
                UserType          = Get-SafeProperty $u 'UserType'
                AccountEnabled    = Get-SafeProperty $u 'AccountEnabled'
                Department        = Get-SafeProperty $u 'Department'
                JobTitle          = Get-SafeProperty $u 'JobTitle'
                LastSuccessfulSignIn = $lastSuccess
                DaysInactive      = Get-DaysSince $lastSuccess
                NeverSignedIn     = [bool](-not $lastSuccess)
                LicenseCount      = $licenses.Count
                CreatedDateTime   = Get-SafeProperty $u 'CreatedDateTime'
                UserId            = Get-SafeProperty $u 'Id'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'identity.guest-users'
    Name     = 'Guest Users'
    Category = 'Identity'
    Synopsis = 'Guest accounts with invitation state and last sign-in.'
    Service  = 'Graph'
    Scopes   = @('User.Read.All', 'AuditLog.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Guest User Report / Guest Users Last Logon Time Report'
    Parameters = @(
        @{ Name = 'StaleDays'; Type = 'Int'; Default = 0
           Help = 'If greater than zero, return only guests inactive for this many days.' }
    )
    Execute = {
        param($P)

        $props = @('Id','DisplayName','UserPrincipalName','Mail','AccountEnabled',
                   'ExternalUserState','ExternalUserStateChangeDateTime','CreatedDateTime','SignInActivity')

        $guests = Get-MgUser -All -Filter "userType eq 'Guest'" -Property $props -ErrorAction Stop

        foreach ($g in $guests) {
            $last = Get-SafeProperty $g 'SignInActivity.LastSuccessfulSignInDateTime'
            if (-not $last) { $last = Get-SafeProperty $g 'SignInActivity.LastSignInDateTime' }
            $days = Get-DaysSince $last

            if ([int]$P.StaleDays -gt 0) {
                # Never-signed-in guests count as stale.
                if ($null -ne $days -and $days -lt [int]$P.StaleDays) { continue }
            }

            [pscustomobject]@{
                DisplayName       = Get-SafeProperty $g 'DisplayName'
                UserPrincipalName = Get-SafeProperty $g 'UserPrincipalName'
                Mail              = Get-SafeProperty $g 'Mail'
                InvitationState   = Get-SafeProperty $g 'ExternalUserState'
                StateChanged      = Get-SafeProperty $g 'ExternalUserStateChangeDateTime'
                AccountEnabled    = Get-SafeProperty $g 'AccountEnabled'
                CreatedDateTime   = Get-SafeProperty $g 'CreatedDateTime'
                LastSignIn        = $last
                DaysInactive      = $days
                NeverSignedIn     = [bool](-not $last)
                UserId            = Get-SafeProperty $g 'Id'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'identity.admin-roles'
    Name     = 'Privileged Role Assignments'
    Category = 'Identity'
    Synopsis = 'Who holds which directory role, with MFA state for each holder.'
    Service  = 'Graph'
    Scopes   = @('RoleManagement.Read.Directory', 'User.Read.All', 'AuditLog.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Office 365 Admin Report'
    Execute = {
        param($P)

        # Build an MFA lookup once rather than per role member.
        $mfaLookup = @{}
        try {
            foreach ($d in (Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop)) {
                $upn = Get-SafeProperty $d 'UserPrincipalName'
                if ($upn) { $mfaLookup[$upn.ToLowerInvariant()] = (Get-SafeProperty $d 'IsMfaRegistered' $false) }
            }
        } catch {
            Write-M365Log -Level Warning -Source 'identity.admin-roles' -Message "MFA registration report unavailable: $($_.Exception.Message)"
        }

        $roles = Get-MgDirectoryRole -All -ErrorAction Stop

        foreach ($role in $roles) {
            $roleName = Get-SafeProperty $role 'DisplayName'
            $members = @()
            try {
                $members = Get-MgDirectoryRoleMember -DirectoryRoleId (Get-SafeProperty $role 'Id') -All -ErrorAction Stop
            } catch {
                Write-M365Log -Level Warning -Source 'identity.admin-roles' -Message "Members of '$roleName': $($_.Exception.Message)"
                continue
            }

            foreach ($m in $members) {
                $ap  = Get-SafeProperty $m 'AdditionalProperties' @{}
                $upn = if ($ap -is [System.Collections.IDictionary] -and $ap.Contains('userPrincipalName')) { $ap['userPrincipalName'] } else { $null }
                $dn  = if ($ap -is [System.Collections.IDictionary] -and $ap.Contains('displayName'))       { $ap['displayName'] }       else { $null }
                $ot  = if ($ap -is [System.Collections.IDictionary] -and $ap.Contains('@odata.type'))       { ($ap['@odata.type'] -replace '#microsoft.graph.','') } else { 'unknown' }

                [pscustomobject]@{
                    RoleName          = $roleName
                    MemberType        = $ot
                    DisplayName       = $dn
                    UserPrincipalName = $upn
                    MfaRegistered     = if ($upn -and $mfaLookup.ContainsKey($upn.ToLowerInvariant())) { $mfaLookup[$upn.ToLowerInvariant()] } else { $null }
                    ObjectId          = Get-SafeProperty $m 'Id'
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'identity.devices'
    Name     = 'Registered Devices'
    Category = 'Identity'
    Synopsis = 'Entra-registered and joined devices with compliance and staleness.'
    Service  = 'Graph'
    Scopes   = @('Device.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Azure AD Devices Report (retired AzureAD module Get-AzureADDevice)'
    Parameters = @(
        @{ Name = 'StaleDays'; Type = 'Int'; Default = 0
           Help = 'If greater than zero, return only devices not seen for this many days.' }
        @{ Name = 'OnlyNonCompliant'; Type = 'Bool'; Default = $false
           Help = 'Return only devices flagged non-compliant.' }
    )
    Execute = {
        param($P)

        $devices = Get-MgDevice -All -ErrorAction Stop

        foreach ($d in $devices) {
            $lastSeen = Get-SafeProperty $d 'ApproximateLastSignInDateTime'
            $days = Get-DaysSince $lastSeen
            $compliant = Get-SafeProperty $d 'IsCompliant'

            if ([int]$P.StaleDays -gt 0 -and $null -ne $days -and $days -lt [int]$P.StaleDays) { continue }
            if ($P.OnlyNonCompliant -and $compliant -ne $false) { continue }

            [pscustomobject]@{
                DisplayName      = Get-SafeProperty $d 'DisplayName'
                OperatingSystem  = Get-SafeProperty $d 'OperatingSystem'
                OSVersion        = Get-SafeProperty $d 'OperatingSystemVersion'
                TrustType        = Get-SafeProperty $d 'TrustType'
                IsCompliant      = $compliant
                IsManaged        = Get-SafeProperty $d 'IsManaged'
                AccountEnabled   = Get-SafeProperty $d 'AccountEnabled'
                LastSignIn       = $lastSeen
                DaysSinceSeen    = $days
                RegisteredDate   = Get-SafeProperty $d 'RegistrationDateTime'
                DeviceId         = Get-SafeProperty $d 'DeviceId'
                ObjectId         = Get-SafeProperty $d 'Id'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'identity.user-membership'
    Name     = 'User Direct Memberships'
    Category = 'Identity'
    Synopsis = 'Groups, directory roles and administrative units a user belongs to.'
    Service  = 'Graph'
    Scopes   = @('User.Read.All', 'Directory.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Get Microsoft Users Direct Membership (Groups, Directory Roles, and AUs)'
    Parameters = @(
        @{ Name = 'UserPrincipalName'; Type = 'String'; Required = $true
           Help = 'UPN of the user to inspect.' }
    )
    Execute = {
        param($P)

        $upn = $P.UserPrincipalName.Trim()
        $memberships = Get-MgUserMemberOf -UserId $upn -All -ErrorAction Stop

        foreach ($m in $memberships) {
            $ap = Get-SafeProperty $m 'AdditionalProperties' @{}
            $type = if ($ap -is [System.Collections.IDictionary] -and $ap.Contains('@odata.type')) {
                        ($ap['@odata.type'] -replace '#microsoft.graph.','')
                    } else { 'unknown' }

            [pscustomobject]@{
                UserPrincipalName = $upn
                ObjectType        = $type
                DisplayName       = if ($ap.Contains('displayName'))  { $ap['displayName'] }  else { $null }
                Description       = if ($ap.Contains('description'))  { $ap['description'] }  else { $null }
                Mail              = if ($ap.Contains('mail'))         { $ap['mail'] }         else { $null }
                ObjectId          = Get-SafeProperty $m 'Id'
            }
        }
    }
}
