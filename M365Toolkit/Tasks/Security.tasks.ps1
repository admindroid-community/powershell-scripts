<#
    Security and incident-response tasks.
#>

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'security.risky-users'
    Name     = 'Risky Users'
    Category = 'Security'
    Synopsis = 'Entra ID Protection risk state per user.'
    Service  = 'Graph'
    Scopes   = @('IdentityRiskyUser.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Get Risky Users in Entra'
    Notes    = 'Requires Entra ID P2 for full risk detail.'
    Parameters = @(
        @{ Name = 'MinimumRiskLevel'; Type = 'Choice'; Default = 'low'
           Options = @('low','medium','high'); Help = 'Lowest risk level to include.' }
        @{ Name = 'ExcludeDismissed'; Type = 'Bool'; Default = $true
           Help = 'Hide users whose risk has been dismissed or remediated.' }
    )
    Execute = {
        param($P)

        $rank = @{ none = 0; low = 1; medium = 2; high = 3 }
        $floor = $rank[$P.MinimumRiskLevel]

        $users = Get-MgRiskyUser -All -ErrorAction Stop

        foreach ($u in $users) {
            $level = (Get-SafeProperty $u 'RiskLevel' 'none').ToString().ToLowerInvariant()
            $state = (Get-SafeProperty $u 'RiskState' '').ToString().ToLowerInvariant()

            if (-not $rank.ContainsKey($level)) { continue }
            if ($rank[$level] -lt $floor) { continue }
            if ($P.ExcludeDismissed -and $state -in @('dismissed','remediated')) { continue }

            [pscustomobject]@{
                DisplayName       = Get-SafeProperty $u 'UserDisplayName'
                UserPrincipalName = Get-SafeProperty $u 'UserPrincipalName'
                RiskLevel         = Get-SafeProperty $u 'RiskLevel'
                RiskState         = Get-SafeProperty $u 'RiskState'
                RiskDetail        = Get-SafeProperty $u 'RiskDetail'
                LastUpdated       = Get-SafeProperty $u 'RiskLastUpdatedDateTime'
                DaysSinceUpdate   = Get-DaysSince (Get-SafeProperty $u 'RiskLastUpdatedDateTime')
                UserId            = Get-SafeProperty $u 'Id'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'security.failed-signins'
    Name     = 'Failed Sign-in Attempts'
    Category = 'Security'
    Synopsis = 'Failed interactive sign-ins, grouped by user and failure reason.'
    Service  = 'Graph'
    Scopes   = @('AuditLog.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Audit Failed Signin Attempts in Entra'
    Parameters = @(
        @{ Name = 'Days'; Type = 'Int'; Default = 7
           Help = 'Look-back window. Entra retains sign-in logs for 7-30 days depending on licence.' }
        @{ Name = 'MinimumAttempts'; Type = 'Int'; Default = 1
           Help = 'Only report users with at least this many failures.' }
        @{ Name = 'UserPrincipalName'; Type = 'String'
           Help = 'Optional: restrict to a single user.' }
    )
    Execute = {
        param($P)

        $since = (Get-Date).AddDays(-1 * [int]$P.Days).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

        # errorCode 0 means success, so anything non-zero is a failure.
        $filter = "createdDateTime ge $since and status/errorCode ne 0"
        if ($P.UserPrincipalName) {
            $filter += " and userPrincipalName eq '$($P.UserPrincipalName.Trim())'"
        }

        $signIns = Get-MgAuditLogSignIn -Filter $filter -All -ErrorAction Stop

        $signIns |
            Group-Object -Property UserPrincipalName |
            Where-Object { $_.Count -ge [int]$P.MinimumAttempts } |
            ForEach-Object {
                $group = $_
                $reasons = $group.Group |
                    ForEach-Object { Get-SafeProperty $_ 'Status.FailureReason' } |
                    Where-Object { $_ } | Group-Object | Sort-Object Count -Descending

                $ips = $group.Group |
                    ForEach-Object { Get-SafeProperty $_ 'IPAddress' } |
                    Where-Object { $_ } | Select-Object -Unique

                $countries = $group.Group |
                    ForEach-Object { Get-SafeProperty $_ 'Location.CountryOrRegion' } |
                    Where-Object { $_ } | Select-Object -Unique

                [pscustomobject]@{
                    UserPrincipalName = $group.Name
                    FailureCount      = $group.Count
                    TopFailureReason  = if ($reasons) { $reasons[0].Name } else { $null }
                    DistinctReasons   = ($reasons | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join '; '
                    DistinctIPs       = $ips.Count
                    IPAddresses       = ($ips | Select-Object -First 10) -join '; '
                    Countries         = ($countries | Select-Object -First 10) -join '; '
                    FirstAttempt      = ($group.Group | Sort-Object CreatedDateTime | Select-Object -First 1).CreatedDateTime
                    LastAttempt       = ($group.Group | Sort-Object CreatedDateTime | Select-Object -Last 1).CreatedDateTime
                }
            } | Sort-Object FailureCount -Descending
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'security.external-forwarding'
    Name     = 'External Email Forwarding'
    Category = 'Security'
    Synopsis = 'Mailboxes and inbox rules forwarding mail outside the tenant.'
    Service  = 'ExchangeOnline'
    Risk     = 'ReadOnly'
    Replaces = 'Find Inbox Rules that Forwards Email Externally / Office 365 Email Forwarding Report'
    Notes    = 'Covers both mailbox-level forwarding and per-rule forwarding, which are configured separately and are a common BEC exfiltration path.'
    Parameters = @(
        @{ Name = 'IncludeInboxRules'; Type = 'Bool'; Default = $true
           Help = 'Also scan inbox rules. Slower - one call per mailbox.' }
        @{ Name = 'MailboxFilter'; Type = 'String'
           Help = 'Optional: limit to a single mailbox identity.' }
    )
    Execute = {
        param($P)

        $accepted = Get-M365AcceptedDomain

        $mbxParams = @{ ResultSize = 'Unlimited'; ErrorAction = 'Stop' }
        if ($P.MailboxFilter) { $mbxParams.Remove('ResultSize'); $mbxParams['Identity'] = $P.MailboxFilter }

        $mailboxes = Get-Mailbox @mbxParams

        foreach ($mbx in $mailboxes) {
            $upn = Get-SafeProperty $mbx 'UserPrincipalName'

            # --- mailbox-level forwarding ---
            $fwdSmtp = Get-SafeProperty $mbx 'ForwardingSmtpAddress'
            $fwdAddr = Get-SafeProperty $mbx 'ForwardingAddress'

            if ($fwdSmtp) {
                $clean = ($fwdSmtp -replace '^smtp:', '').Trim()
                if (Test-ExternalDomain -Address $clean -AcceptedDomains $accepted) {
                    [pscustomobject]@{
                        Mailbox           = Get-SafeProperty $mbx 'DisplayName'
                        UserPrincipalName = $upn
                        FindingType       = 'MailboxForwarding'
                        RuleName          = $null
                        ForwardsTo        = $clean
                        DeliverAndForward = Get-SafeProperty $mbx 'DeliverToMailboxAndForward'
                        Enabled           = $true
                        Severity          = 'High'
                    }
                }
            }

            if ($fwdAddr) {
                [pscustomobject]@{
                    Mailbox           = Get-SafeProperty $mbx 'DisplayName'
                    UserPrincipalName = $upn
                    FindingType       = 'MailboxForwardingAddress'
                    RuleName          = $null
                    ForwardsTo        = $fwdAddr.ToString()
                    DeliverAndForward = Get-SafeProperty $mbx 'DeliverToMailboxAndForward'
                    Enabled           = $true
                    Severity          = 'Medium'
                }
            }

            # --- inbox rules ---
            if (-not $P.IncludeInboxRules) { continue }

            try {
                $rules = Get-InboxRule -Mailbox $upn -ErrorAction Stop
            } catch {
                Write-M365Log -Level Warning -Source 'security.external-forwarding' -Message "Inbox rules for $upn : $($_.Exception.Message)"
                continue
            }

            foreach ($rule in $rules) {
                $targets = @()
                foreach ($prop in 'ForwardTo', 'ForwardAsAttachmentTo', 'RedirectTo') {
                    $vals = Get-SafeProperty $rule $prop @()
                    foreach ($v in @($vals)) {
                        if ($null -eq $v) { continue }
                        $text = $v.ToString()
                        # Exchange formats these as "Name [SMTP:addr]"
                        if ($text -match 'SMTP:([^\]\s]+)') { $targets += $Matches[1] }
                        elseif ($text -match '[\w.+-]+@[\w.-]+')  { $targets += $Matches[0] }
                    }
                }

                $external = @($targets | Where-Object { Test-ExternalDomain -Address $_ -AcceptedDomains $accepted } | Select-Object -Unique)
                if ($external.Count -eq 0) { continue }

                [pscustomobject]@{
                    Mailbox           = Get-SafeProperty $mbx 'DisplayName'
                    UserPrincipalName = $upn
                    FindingType       = 'InboxRule'
                    RuleName          = Get-SafeProperty $rule 'Name'
                    ForwardsTo        = ($external -join '; ')
                    DeliverAndForward = $null
                    Enabled           = Get-SafeProperty $rule 'Enabled'
                    Severity          = 'High'
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'security.block-external-forwarding'
    Name     = 'Block External Forwarding'
    Category = 'Security'
    Synopsis = 'Removes external mailbox forwarding and disables offending inbox rules.'
    Service  = 'ExchangeOnline'
    Risk     = 'Destructive'
    Replaces = 'Identify and Block External Email Forwarding in Exchange Online / Remove Email Forwarding'
    Notes    = 'Run security.external-forwarding first and review. Supports WhatIf.'
    Parameters = @(
        @{ Name = 'UserPrincipalName'; Type = 'String'
           Help = 'Restrict to one mailbox. Leave blank to remediate every finding tenant-wide.' }
        @{ Name = 'DisableInboxRules'; Type = 'Bool'; Default = $true
           Help = 'Disable (not delete) inbox rules that forward externally.' }
        @{ Name = 'WhatIf'; Type = 'Bool'; Default = $true
           Help = 'Report what would change without changing it. Set false to apply.' }
    )
    Execute = {
        param($P)

        $accepted = Get-M365AcceptedDomain
        $dryRun   = [bool]$P.WhatIf

        $mbxParams = @{ ResultSize = 'Unlimited'; ErrorAction = 'Stop' }
        if ($P.UserPrincipalName) { $mbxParams.Remove('ResultSize'); $mbxParams['Identity'] = $P.UserPrincipalName }

        foreach ($mbx in (Get-Mailbox @mbxParams)) {
            $upn = Get-SafeProperty $mbx 'UserPrincipalName'
            $fwdSmtp = Get-SafeProperty $mbx 'ForwardingSmtpAddress'

            if ($fwdSmtp) {
                $clean = ($fwdSmtp -replace '^smtp:', '').Trim()
                if (Test-ExternalDomain -Address $clean -AcceptedDomains $accepted) {
                    $status = 'WouldRemove'
                    if (-not $dryRun) {
                        try {
                            Set-Mailbox -Identity $upn -ForwardingSmtpAddress $null -ErrorAction Stop
                            $status = 'Removed'
                        } catch {
                            $status = "Failed: $($_.Exception.Message)"
                        }
                    }
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        Action            = 'MailboxForwarding'
                        Target            = $clean
                        RuleName          = $null
                        Status            = $status
                    }
                }
            }

            if (-not $P.DisableInboxRules) { continue }

            try { $rules = Get-InboxRule -Mailbox $upn -ErrorAction Stop } catch { continue }

            foreach ($rule in $rules) {
                $targets = @()
                foreach ($prop in 'ForwardTo', 'ForwardAsAttachmentTo', 'RedirectTo') {
                    foreach ($v in @(Get-SafeProperty $rule $prop @())) {
                        if ($null -eq $v) { continue }
                        $text = $v.ToString()
                        if ($text -match 'SMTP:([^\]\s]+)') { $targets += $Matches[1] }
                        elseif ($text -match '[\w.+-]+@[\w.-]+')  { $targets += $Matches[0] }
                    }
                }
                $external = @($targets | Where-Object { Test-ExternalDomain -Address $_ -AcceptedDomains $accepted } | Select-Object -Unique)
                if ($external.Count -eq 0) { continue }
                if (-not (Get-SafeProperty $rule 'Enabled' $false)) { continue }

                $status = 'WouldDisable'
                if (-not $dryRun) {
                    try {
                        Disable-InboxRule -Identity (Get-SafeProperty $rule 'Identity') -Mailbox $upn -Confirm:$false -ErrorAction Stop
                        $status = 'Disabled'
                    } catch {
                        $status = "Failed: $($_.Exception.Message)"
                    }
                }

                [pscustomobject]@{
                    UserPrincipalName = $upn
                    Action            = 'InboxRule'
                    Target            = ($external -join '; ')
                    RuleName          = Get-SafeProperty $rule 'Name'
                    Status            = $status
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'security.compromised-remediation'
    Name     = 'Compromised Account Remediation'
    Category = 'Security'
    Synopsis = 'Revokes sessions, optionally blocks sign-in and strips forwarding for an account.'
    Service  = @('Graph', 'ExchangeOnline')
    Scopes   = @('User.ReadWrite.All', 'Directory.AccessAsUser.All')
    Risk     = 'Destructive'
    Replaces = 'Automate Compromised Account Remediation'
    Notes    = 'Session revocation forces re-authentication everywhere. Runs in WhatIf by default.'
    Parameters = @(
        @{ Name = 'UserPrincipalName'; Type = 'String'; Required = $true
           Help = 'Account to remediate.' }
        @{ Name = 'RevokeSessions'; Type = 'Bool'; Default = $true
           Help = 'Invalidate all refresh tokens.' }
        @{ Name = 'BlockSignIn'; Type = 'Bool'; Default = $true
           Help = 'Disable the account.' }
        @{ Name = 'RemoveForwarding'; Type = 'Bool'; Default = $true
           Help = 'Clear mailbox forwarding and disable all inbox rules.' }
        @{ Name = 'WhatIf'; Type = 'Bool'; Default = $true
           Help = 'Report intended actions without applying them. Set false to apply.' }
    )
    Execute = {
        param($P)

        $upn    = $P.UserPrincipalName.Trim()
        $dryRun = [bool]$P.WhatIf

        function New-Step($action, $status, $detail) {
            [pscustomobject]@{
                UserPrincipalName = $upn
                Action            = $action
                Status            = $status
                Detail            = $detail
                Timestamp         = Get-Date
            }
        }

        $user = Get-MgUser -UserId $upn -Property 'Id','DisplayName','AccountEnabled' -ErrorAction Stop
        $userId = Get-SafeProperty $user 'Id'
        New-Step 'Resolve' 'OK' "Found $(Get-SafeProperty $user 'DisplayName') ($userId)"

        if ($P.RevokeSessions) {
            if ($dryRun) { New-Step 'RevokeSessions' 'WouldRun' 'Invalidate all refresh tokens' }
            else {
                try {
                    Revoke-MgUserSignInSession -UserId $userId -ErrorAction Stop | Out-Null
                    New-Step 'RevokeSessions' 'Done' 'All refresh tokens invalidated'
                } catch { New-Step 'RevokeSessions' 'Failed' $_.Exception.Message }
            }
        }

        if ($P.BlockSignIn) {
            if ($dryRun) { New-Step 'BlockSignIn' 'WouldRun' 'Set AccountEnabled = false' }
            else {
                try {
                    Update-MgUser -UserId $userId -AccountEnabled:$false -ErrorAction Stop
                    New-Step 'BlockSignIn' 'Done' 'Account disabled'
                } catch { New-Step 'BlockSignIn' 'Failed' $_.Exception.Message }
            }
        }

        if ($P.RemoveForwarding) {
            if ($dryRun) { New-Step 'RemoveForwarding' 'WouldRun' 'Clear forwarding and disable inbox rules' }
            else {
                try {
                    Set-Mailbox -Identity $upn -ForwardingSmtpAddress $null -ForwardingAddress $null -ErrorAction Stop
                    New-Step 'RemoveForwarding' 'Done' 'Mailbox forwarding cleared'
                } catch { New-Step 'RemoveForwarding' 'Failed' $_.Exception.Message }

                try {
                    $rules = Get-InboxRule -Mailbox $upn -ErrorAction Stop
                    foreach ($r in $rules) {
                        if (-not (Get-SafeProperty $r 'Enabled' $false)) { continue }
                        try {
                            Disable-InboxRule -Identity (Get-SafeProperty $r 'Identity') -Mailbox $upn -Confirm:$false -ErrorAction Stop
                            New-Step 'DisableInboxRule' 'Done' (Get-SafeProperty $r 'Name')
                        } catch { New-Step 'DisableInboxRule' 'Failed' "$(Get-SafeProperty $r 'Name'): $($_.Exception.Message)" }
                    }
                } catch { New-Step 'DisableInboxRule' 'Failed' $_.Exception.Message }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'security.app-credential-expiry'
    Name     = 'App Credential Expiry'
    Category = 'Security'
    Synopsis = 'App registration secrets and certificates that are expiring or expired.'
    Service  = 'Graph'
    Scopes   = @('Application.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Notify Entra App Credential Expiry / App Registrations with Expiring Certificates & Client Secrets'
    Parameters = @(
        @{ Name = 'DaysUntilExpiry'; Type = 'Int'; Default = 30
           Help = 'Warn on credentials expiring within this many days.' }
        @{ Name = 'IncludeExpired'; Type = 'Bool'; Default = $true
           Help = 'Also list credentials that have already expired.' }
    )
    Execute = {
        param($P)

        $horizon = (Get-Date).AddDays([int]$P.DaysUntilExpiry)
        $apps = Get-MgApplication -All -ErrorAction Stop

        foreach ($app in $apps) {
            $appName = Get-SafeProperty $app 'DisplayName'
            $appId   = Get-SafeProperty $app 'AppId'

            $creds = @()
            foreach ($pw in @(Get-SafeProperty $app 'PasswordCredentials' @())) {
                $creds += [pscustomobject]@{ Kind = 'ClientSecret'; Obj = $pw }
            }
            foreach ($cert in @(Get-SafeProperty $app 'KeyCredentials' @())) {
                $creds += [pscustomobject]@{ Kind = 'Certificate'; Obj = $cert }
            }

            foreach ($c in $creds) {
                $end = Get-SafeProperty $c.Obj 'EndDateTime'
                if (-not $end) { continue }

                $endDt = [datetime]$end
                $daysLeft = [int]($endDt - (Get-Date)).TotalDays
                $expired = $endDt -lt (Get-Date)

                if ($expired -and -not $P.IncludeExpired) { continue }
                if (-not $expired -and $endDt -gt $horizon) { continue }

                [pscustomobject]@{
                    ApplicationName = $appName
                    ApplicationId   = $appId
                    CredentialType  = $c.Kind
                    CredentialName  = Get-SafeProperty $c.Obj 'DisplayName'
                    StartDate       = Get-SafeProperty $c.Obj 'StartDateTime'
                    EndDate         = $endDt
                    DaysRemaining   = $daysLeft
                    Status          = if ($expired) { 'Expired' } elseif ($daysLeft -le 7) { 'Critical' } else { 'Expiring' }
                    ObjectId        = Get-SafeProperty $app 'Id'
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'security.enterprise-app-owners'
    Name     = 'Enterprise Apps and Owners'
    Category = 'Security'
    Synopsis = 'Service principals with their owners and assigned app roles.'
    Service  = 'Graph'
    Scopes   = @('Application.Read.All', 'Directory.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Get All Enterprise Apps and their Owners'
    Parameters = @(
        @{ Name = 'OnlyOwnerless'; Type = 'Bool'; Default = $false
           Help = 'Return only applications with no assigned owner.' }
    )
    Execute = {
        param($P)

        $sps = Get-MgServicePrincipal -All -ErrorAction Stop

        foreach ($sp in $sps) {
            $spId = Get-SafeProperty $sp 'Id'

            $owners = @()
            try {
                $owners = Get-MgServicePrincipalOwner -ServicePrincipalId $spId -All -ErrorAction Stop
            } catch { }

            $ownerNames = foreach ($o in $owners) {
                $ap = Get-SafeProperty $o 'AdditionalProperties' @{}
                if ($ap -is [System.Collections.IDictionary] -and $ap.Contains('userPrincipalName')) { $ap['userPrincipalName'] }
                elseif ($ap -is [System.Collections.IDictionary] -and $ap.Contains('displayName'))    { $ap['displayName'] }
            }
            $ownerNames = @($ownerNames | Where-Object { $_ })

            if ($P.OnlyOwnerless -and $ownerNames.Count -gt 0) { continue }

            [pscustomobject]@{
                DisplayName        = Get-SafeProperty $sp 'DisplayName'
                AppId              = Get-SafeProperty $sp 'AppId'
                ServicePrincipalType = Get-SafeProperty $sp 'ServicePrincipalType'
                AccountEnabled     = Get-SafeProperty $sp 'AccountEnabled'
                PublisherName      = Get-SafeProperty $sp 'PublisherName'
                SignInAudience     = Get-SafeProperty $sp 'SignInAudience'
                OwnerCount         = $ownerNames.Count
                Owners             = ($ownerNames -join '; ')
                Homepage           = Get-SafeProperty $sp 'Homepage'
                ObjectId           = $spId
            }
        }
    }
}
