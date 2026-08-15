<#
    Joiner / mover / leaver and group lifecycle tasks.
#>

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'lifecycle.offboard-user'
    Name     = 'Offboard User'
    Category = 'Lifecycle'
    Synopsis = 'Runs the standard leaver sequence and reports each step.'
    Service  = @('Graph', 'ExchangeOnline')
    Scopes   = @('User.ReadWrite.All', 'Directory.ReadWrite.All', 'Group.ReadWrite.All')
    Risk     = 'Destructive'
    Replaces = 'Automate M365 User Offboarding'
    Notes    = 'Defaults to WhatIf. Every step reports individually, so a partial failure is visible rather than silent.'
    Parameters = @(
        @{ Name = 'UserPrincipalName'; Type = 'String'; Required = $true
           Help = 'Leaver account.' }
        @{ Name = 'BlockSignIn'; Type = 'Bool'; Default = $true
           Help = 'Disable the account.' }
        @{ Name = 'RevokeSessions'; Type = 'Bool'; Default = $true
           Help = 'Invalidate refresh tokens.' }
        @{ Name = 'ConvertToShared'; Type = 'Bool'; Default = $true
           Help = 'Convert the mailbox to shared so it stays accessible without a licence.' }
        @{ Name = 'DelegateMailboxTo'; Type = 'String'
           Help = 'Optional: UPN to grant Full Access over the mailbox.' }
        @{ Name = 'SetAutoReply'; Type = 'String'
           Help = 'Optional: automatic reply text to set for external senders.' }
        @{ Name = 'RemoveFromGroups'; Type = 'Bool'; Default = $true
           Help = 'Remove from all cloud distribution and M365 groups.' }
        @{ Name = 'RemoveLicenses'; Type = 'Bool'; Default = $true
           Help = 'Strip directly assigned licences.' }
        @{ Name = 'WhatIf'; Type = 'Bool'; Default = $true
           Help = 'Report the plan without applying it. Set false to execute.' }
    )
    Execute = {
        param($P)

        $upn    = $P.UserPrincipalName.Trim()
        $dryRun = [bool]$P.WhatIf
        $seq    = 0

        function New-Step($action, $status, $detail) {
            $script:__seq = if ($null -eq $script:__seq) { 1 } else { $script:__seq + 1 }
            [pscustomobject]@{
                Step              = $script:__seq
                UserPrincipalName = $upn
                Action            = $action
                Status            = $status
                Detail            = $detail
                Timestamp         = Get-Date
            }
        }
        $script:__seq = 0

        $user = Get-MgUser -UserId $upn -Property 'Id','DisplayName','AccountEnabled','AssignedLicenses' -ErrorAction Stop
        $userId = Get-SafeProperty $user 'Id'
        New-Step 'Resolve' 'OK' "$(Get-SafeProperty $user 'DisplayName') ($userId)"

        # --- block sign-in ---
        if ($P.BlockSignIn) {
            if ($dryRun) { New-Step 'BlockSignIn' 'WouldRun' 'AccountEnabled = false' }
            else {
                try { Update-MgUser -UserId $userId -AccountEnabled:$false -ErrorAction Stop
                      New-Step 'BlockSignIn' 'Done' 'Account disabled' }
                catch { New-Step 'BlockSignIn' 'Failed' $_.Exception.Message }
            }
        }

        # --- revoke sessions ---
        if ($P.RevokeSessions) {
            if ($dryRun) { New-Step 'RevokeSessions' 'WouldRun' 'Invalidate refresh tokens' }
            else {
                try { Revoke-MgUserSignInSession -UserId $userId -ErrorAction Stop | Out-Null
                      New-Step 'RevokeSessions' 'Done' 'Tokens invalidated' }
                catch { New-Step 'RevokeSessions' 'Failed' $_.Exception.Message }
            }
        }

        # --- auto reply (before conversion, while still a user mailbox) ---
        if ($P.SetAutoReply) {
            if ($dryRun) { New-Step 'SetAutoReply' 'WouldRun' $P.SetAutoReply }
            else {
                try {
                    Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled `
                        -InternalMessage $P.SetAutoReply -ExternalMessage $P.SetAutoReply -ErrorAction Stop
                    New-Step 'SetAutoReply' 'Done' 'Automatic replies enabled'
                } catch { New-Step 'SetAutoReply' 'Failed' $_.Exception.Message }
            }
        }

        # --- delegate access ---
        if ($P.DelegateMailboxTo) {
            if ($dryRun) { New-Step 'DelegateMailbox' 'WouldRun' "Full Access for $($P.DelegateMailboxTo)" }
            else {
                try {
                    Add-MailboxPermission -Identity $upn -User $P.DelegateMailboxTo.Trim() `
                        -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                    New-Step 'DelegateMailbox' 'Done' "Full Access granted to $($P.DelegateMailboxTo)"
                } catch { New-Step 'DelegateMailbox' 'Failed' $_.Exception.Message }
            }
        }

        # --- group removal ---
        if ($P.RemoveFromGroups) {
            try {
                $groups = Get-MgUserMemberOf -UserId $userId -All -ErrorAction Stop
                foreach ($g in $groups) {
                    $ap = Get-SafeProperty $g 'AdditionalProperties' @{}
                    $type = if ($ap -is [System.Collections.IDictionary] -and $ap.Contains('@odata.type')) { $ap['@odata.type'] } else { '' }
                    if ($type -notlike '*group*') { continue }   # skip directory roles / AUs
                    $name = if ($ap.Contains('displayName')) { $ap['displayName'] } else { (Get-SafeProperty $g 'Id') }

                    if ($dryRun) { New-Step 'RemoveFromGroup' 'WouldRun' $name; continue }
                    try {
                        Remove-MgGroupMemberByRef -GroupId (Get-SafeProperty $g 'Id') -DirectoryObjectId $userId -ErrorAction Stop
                        New-Step 'RemoveFromGroup' 'Done' $name
                    } catch {
                        New-Step 'RemoveFromGroup' 'Failed' "$name : $($_.Exception.Message)"
                    }
                }
            } catch { New-Step 'RemoveFromGroup' 'Failed' $_.Exception.Message }
        }

        # --- convert mailbox ---
        if ($P.ConvertToShared) {
            if ($dryRun) { New-Step 'ConvertToShared' 'WouldRun' 'Set mailbox type = Shared' }
            else {
                try { Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
                      New-Step 'ConvertToShared' 'Done' 'Mailbox converted to shared' }
                catch { New-Step 'ConvertToShared' 'Failed' $_.Exception.Message }
            }
        }

        # --- licences last, so the mailbox conversion has taken effect ---
        if ($P.RemoveLicenses) {
            $assigned = @(Get-SafeProperty $user 'AssignedLicenses' @())
            if ($assigned.Count -eq 0) {
                New-Step 'RemoveLicenses' 'Skipped' 'No directly assigned licences'
            } elseif ($dryRun) {
                New-Step 'RemoveLicenses' 'WouldRun' "$($assigned.Count) licence(s)"
            } else {
                try {
                    $skuIds = @($assigned | ForEach-Object { Get-SafeProperty $_ 'SkuId' } | Where-Object { $_ })
                    Set-MgUserLicense -UserId $userId -AddLicenses @() -RemoveLicenses $skuIds -ErrorAction Stop | Out-Null
                    New-Step 'RemoveLicenses' 'Done' "$($skuIds.Count) licence(s) removed"
                } catch { New-Step 'RemoveLicenses' 'Failed' $_.Exception.Message }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'lifecycle.convert-to-shared'
    Name     = 'Convert Mailboxes to Shared'
    Category = 'Lifecycle'
    Synopsis = 'Bulk-converts user mailboxes to shared and optionally strips licences.'
    Service  = 'ExchangeOnline'
    Risk     = 'Write'
    Replaces = 'Convert User Mailboxes to Shared Mailboxes in Bulk'
    Parameters = @(
        @{ Name = 'UserPrincipalName'; Type = 'String'; Required = $true
           Help = 'One or more UPNs, comma separated.' }
        @{ Name = 'WhatIf'; Type = 'Bool'; Default = $true
           Help = 'Report without applying. Set false to convert.' }
    )
    Execute = {
        param($P)

        $upns = @($P.UserPrincipalName -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

        foreach ($upn in $upns) {
            try {
                $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop
            } catch {
                [pscustomobject]@{ UserPrincipalName = $upn; CurrentType = $null; Status = "Failed: $($_.Exception.Message)" }
                continue
            }

            $current = Get-SafeProperty $mbx 'RecipientTypeDetails'
            if ($current -eq 'SharedMailbox') {
                [pscustomobject]@{ UserPrincipalName = $upn; CurrentType = $current; Status = 'Already shared' }
                continue
            }

            if ($P.WhatIf) {
                [pscustomobject]@{ UserPrincipalName = $upn; CurrentType = $current; Status = 'WouldConvert' }
                continue
            }

            try {
                Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
                [pscustomobject]@{ UserPrincipalName = $upn; CurrentType = $current; Status = 'Converted' }
            } catch {
                [pscustomobject]@{ UserPrincipalName = $upn; CurrentType = $current; Status = "Failed: $($_.Exception.Message)" }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'lifecycle.m365-groups'
    Name     = 'Microsoft 365 Groups'
    Category = 'Lifecycle'
    Synopsis = 'M365 groups with owners, membership counts and orphan detection.'
    Service  = 'Graph'
    Scopes   = @('Group.Read.All', 'Directory.Read.All')
    Risk     = 'ReadOnly'
    Replaces = 'Microsoft 365 Group Report (retired MSOnline dependency) / Find Orphaned Teams'
    Notes    = 'An ownerless group is a governance gap: nobody can approve access or action its lifecycle.'
    Parameters = @(
        @{ Name = 'OnlyOrphaned'; Type = 'Bool'; Default = $false
           Help = 'Only groups with no owner.' }
        @{ Name = 'OnlyTeams'; Type = 'Bool'; Default = $false
           Help = 'Only groups that are Teams-enabled.' }
    )
    Execute = {
        param($P)

        $groups = Get-MgGroup -All -Property 'Id','DisplayName','Mail','GroupTypes','Description',
                                            'Visibility','CreatedDateTime','ResourceProvisioningOptions' -ErrorAction Stop

        foreach ($g in $groups) {
            $types = @(Get-SafeProperty $g 'GroupTypes' @())
            if ('Unified' -notin $types) { continue }

            $provisioning = @(Get-SafeProperty $g 'ResourceProvisioningOptions' @())
            $isTeam = 'Team' -in $provisioning
            if ($P.OnlyTeams -and -not $isTeam) { continue }

            $gid = Get-SafeProperty $g 'Id'

            $owners = @(); $memberCount = $null
            try { $owners = Get-MgGroupOwner -GroupId $gid -All -ErrorAction Stop } catch { }
            try { $memberCount = @(Get-MgGroupMember -GroupId $gid -All -ErrorAction Stop).Count } catch { }

            if ($P.OnlyOrphaned -and $owners.Count -gt 0) { continue }

            $ownerNames = foreach ($o in $owners) {
                $ap = Get-SafeProperty $o 'AdditionalProperties' @{}
                if ($ap -is [System.Collections.IDictionary] -and $ap.Contains('userPrincipalName')) { $ap['userPrincipalName'] }
                elseif ($ap -is [System.Collections.IDictionary] -and $ap.Contains('displayName'))    { $ap['displayName'] }
            }

            [pscustomobject]@{
                DisplayName     = Get-SafeProperty $g 'DisplayName'
                Mail            = Get-SafeProperty $g 'Mail'
                IsTeam          = $isTeam
                Visibility      = Get-SafeProperty $g 'Visibility'
                OwnerCount      = $owners.Count
                Owners          = (@($ownerNames | Where-Object { $_ }) -join '; ')
                IsOrphaned      = ($owners.Count -eq 0)
                MemberCount     = $memberCount
                Description     = Get-SafeProperty $g 'Description'
                CreatedDateTime = Get-SafeProperty $g 'CreatedDateTime'
                GroupId         = $gid
            }
        }
    }
}
