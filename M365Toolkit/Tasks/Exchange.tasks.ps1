<#
    Exchange Online tasks.
#>

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'exchange.mailbox-sizes'
    Name     = 'Mailbox Size & Quota'
    Category = 'Exchange'
    Synopsis = 'Mailbox and archive size against quota, with usage percentage.'
    Service  = 'ExchangeOnline'
    Risk     = 'ReadOnly'
    Replaces = 'Mailbox Size Report / Export Mailbox Usage & Quota Report / Archive Mailbox Size Report'
    Parameters = @(
        @{ Name = 'MailboxType'; Type = 'Choice'; Default = 'All'
           Options = @('All','UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox')
           Help = 'Recipient type to include.' }
        @{ Name = 'MinimumUsagePercent'; Type = 'Int'; Default = 0
           Help = 'Only mailboxes above this percentage of quota.' }
        @{ Name = 'IncludeArchive'; Type = 'Bool'; Default = $false
           Help = 'Also report archive mailbox size. Adds a call per mailbox.' }
    )
    Execute = {
        param($P)

        $params = @{ ResultSize = 'Unlimited'; ErrorAction = 'Stop' }
        if ($P.MailboxType -ne 'All') { $params['RecipientTypeDetails'] = $P.MailboxType }

        foreach ($mbx in (Get-Mailbox @params)) {
            $upn = Get-SafeProperty $mbx 'UserPrincipalName'

            try {
                $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop
            } catch {
                Write-M365Log -Level Warning -Source 'exchange.mailbox-sizes' -Message "Stats for $upn : $($_.Exception.Message)"
                continue
            }

            $usedBytes  = ConvertFrom-ExchangeSize (Get-SafeProperty $stats 'TotalItemSize')
            $quotaRaw   = Get-SafeProperty $mbx 'ProhibitSendReceiveQuota'
            $quotaBytes = ConvertFrom-ExchangeSize $quotaRaw

            $usagePct = if ($usedBytes -and $quotaBytes -and $quotaBytes -gt 0) {
                [math]::Round(($usedBytes / $quotaBytes) * 100, 1)
            } else { $null }

            if ([int]$P.MinimumUsagePercent -gt 0) {
                if ($null -eq $usagePct -or $usagePct -lt [int]$P.MinimumUsagePercent) { continue }
            }

            $archiveGB = $null
            if ($P.IncludeArchive -and (Get-SafeProperty $mbx 'ArchiveStatus') -eq 'Active') {
                try {
                    $ast = Get-MailboxStatistics -Identity $upn -Archive -ErrorAction Stop
                    $archiveGB = ConvertTo-FriendlySize (ConvertFrom-ExchangeSize (Get-SafeProperty $ast 'TotalItemSize'))
                } catch { }
            }

            [pscustomobject]@{
                DisplayName       = Get-SafeProperty $mbx 'DisplayName'
                UserPrincipalName = $upn
                MailboxType       = Get-SafeProperty $mbx 'RecipientTypeDetails'
                PrimarySmtpAddress= Get-SafeProperty $mbx 'PrimarySmtpAddress'
                ItemCount         = Get-SafeProperty $stats 'ItemCount'
                SizeGB            = ConvertTo-FriendlySize $usedBytes
                QuotaGB           = ConvertTo-FriendlySize $quotaBytes
                UsagePercent      = $usagePct
                DeletedItemCount  = Get-SafeProperty $stats 'DeletedItemCount'
                ArchiveStatus     = Get-SafeProperty $mbx 'ArchiveStatus'
                ArchiveSizeGB     = $archiveGB
                LitigationHold    = Get-SafeProperty $mbx 'LitigationHoldEnabled'
                LastLogon         = Get-SafeProperty $stats 'LastLogonTime'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'exchange.mailbox-permissions'
    Name     = 'Mailbox Permissions'
    Category = 'Exchange'
    Synopsis = 'Full Access, Send As and Send on Behalf delegations per mailbox.'
    Service  = 'ExchangeOnline'
    Risk     = 'ReadOnly'
    Replaces = 'Office 365 Mailbox Permissions Report / List Mailboxes Users Can Access / Office 365 Shared Mailbox Permission Report'
    Notes    = 'The original relied on a Prerequisites.ps1 that installs the retired MSOnline module.'
    Parameters = @(
        @{ Name = 'MailboxType'; Type = 'Choice'; Default = 'All'
           Options = @('All','UserMailbox','SharedMailbox','RoomMailbox')
           Help = 'Recipient type to inspect.' }
        @{ Name = 'PermissionType'; Type = 'MultiChoice'; Default = @('FullAccess','SendAs','SendOnBehalf')
           Options = @('FullAccess','SendAs','SendOnBehalf')
           Help = 'Which delegation types to report.' }
        @{ Name = 'ExcludeInherited'; Type = 'Bool'; Default = $true
           Help = 'Hide inherited and self permissions - normally noise.' }
    )
    Execute = {
        param($P)

        $wanted = @($P.PermissionType)
        $params = @{ ResultSize = 'Unlimited'; ErrorAction = 'Stop' }
        if ($P.MailboxType -ne 'All') { $params['RecipientTypeDetails'] = $P.MailboxType }

        foreach ($mbx in (Get-Mailbox @params)) {
            $identity = Get-SafeProperty $mbx 'UserPrincipalName'
            $display  = Get-SafeProperty $mbx 'DisplayName'
            $type     = Get-SafeProperty $mbx 'RecipientTypeDetails'

            if ('FullAccess' -in $wanted) {
                try {
                    foreach ($perm in (Get-MailboxPermission -Identity $identity -ErrorAction Stop)) {
                        $user = Get-SafeProperty $perm 'User'
                        if (-not $user) { continue }
                        if ($user -like 'NT AUTHORITY\*' -or $user -like 'S-1-5-*') { continue }
                        if ($P.ExcludeInherited -and (Get-SafeProperty $perm 'IsInherited' $false)) { continue }
                        if ($user -eq $identity) { continue }

                        [pscustomobject]@{
                            Mailbox        = $display
                            MailboxAddress = $identity
                            MailboxType    = $type
                            PermissionType = 'FullAccess'
                            GrantedTo      = $user.ToString()
                            AccessRights   = (@(Get-SafeProperty $perm 'AccessRights' @()) -join ', ')
                            IsInherited    = Get-SafeProperty $perm 'IsInherited'
                        }
                    }
                } catch {
                    Write-M365Log -Level Warning -Source 'exchange.mailbox-permissions' -Message "FullAccess on $identity : $($_.Exception.Message)"
                }
            }

            if ('SendAs' -in $wanted) {
                try {
                    foreach ($perm in (Get-RecipientPermission -Identity $identity -ErrorAction Stop)) {
                        $trustee = Get-SafeProperty $perm 'Trustee'
                        if (-not $trustee -or $trustee -eq 'NT AUTHORITY\SELF') { continue }

                        [pscustomobject]@{
                            Mailbox        = $display
                            MailboxAddress = $identity
                            MailboxType    = $type
                            PermissionType = 'SendAs'
                            GrantedTo      = $trustee.ToString()
                            AccessRights   = (@(Get-SafeProperty $perm 'AccessRights' @()) -join ', ')
                            IsInherited    = $false
                        }
                    }
                } catch {
                    Write-M365Log -Level Warning -Source 'exchange.mailbox-permissions' -Message "SendAs on $identity : $($_.Exception.Message)"
                }
            }

            if ('SendOnBehalf' -in $wanted) {
                foreach ($grantee in @(Get-SafeProperty $mbx 'GrantSendOnBehalfTo' @())) {
                    if (-not $grantee) { continue }
                    [pscustomobject]@{
                        Mailbox        = $display
                        MailboxAddress = $identity
                        MailboxType    = $type
                        PermissionType = 'SendOnBehalf'
                        GrantedTo      = $grantee.ToString()
                        AccessRights   = 'SendOnBehalf'
                        IsInherited    = $false
                    }
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'exchange.shared-mailboxes'
    Name     = 'Shared Mailbox Report'
    Category = 'Exchange'
    Synopsis = 'Shared mailboxes with size, licence state and sign-in status.'
    Service  = 'ExchangeOnline'
    Risk     = 'ReadOnly'
    Replaces = 'Shared Mailbox report / Shared Mailbox Size Report / Find Non-Compliant Shared Mailboxes'
    Notes    = 'A shared mailbox with sign-in enabled is a standing security finding.'
    Execute = {
        param($P)

        foreach ($mbx in (Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)) {
            $upn = Get-SafeProperty $mbx 'UserPrincipalName'

            $sizeGB = $null; $itemCount = $null
            try {
                $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop
                $sizeGB = ConvertTo-FriendlySize (ConvertFrom-ExchangeSize (Get-SafeProperty $stats 'TotalItemSize'))
                $itemCount = Get-SafeProperty $stats 'ItemCount'
            } catch { }

            # Sign-in state comes from Entra, not Exchange.
            $signInBlocked = $null
            if (Get-Command Get-MgUser -ErrorAction SilentlyContinue) {
                try {
                    $u = Get-MgUser -UserId $upn -Property 'AccountEnabled' -ErrorAction Stop
                    $signInBlocked = -not (Get-SafeProperty $u 'AccountEnabled' $true)
                } catch { }
            }

            [pscustomobject]@{
                DisplayName        = Get-SafeProperty $mbx 'DisplayName'
                UserPrincipalName  = $upn
                PrimarySmtpAddress = Get-SafeProperty $mbx 'PrimarySmtpAddress'
                SizeGB             = $sizeGB
                ItemCount          = $itemCount
                SignInBlocked      = $signInBlocked
                ComplianceFlag     = if ($signInBlocked -eq $false) { 'Sign-in NOT blocked' } else { $null }
                ArchiveStatus      = Get-SafeProperty $mbx 'ArchiveStatus'
                LitigationHold     = Get-SafeProperty $mbx 'LitigationHoldEnabled'
                AuditEnabled       = Get-SafeProperty $mbx 'AuditEnabled'
                ForwardingAddress  = Get-SafeProperty $mbx 'ForwardingSmtpAddress'
                WhenCreated        = Get-SafeProperty $mbx 'WhenCreated'
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'exchange.distribution-groups'
    Name     = 'Distribution Groups & Members'
    Category = 'Exchange'
    Synopsis = 'Distribution lists with membership, owners and external members.'
    Service  = 'ExchangeOnline'
    Risk     = 'ReadOnly'
    Replaces = 'Export Distribution Groups / Office 365 Distribution Group Members Report / Find DLs with External Users'
    Parameters = @(
        @{ Name = 'IncludeMembers'; Type = 'Bool'; Default = $true
           Help = 'Expand membership. One row per member when enabled.' }
        @{ Name = 'OnlyWithExternalMembers'; Type = 'Bool'; Default = $false
           Help = 'Only groups containing members outside the tenant.' }
    )
    Execute = {
        param($P)

        $accepted = Get-M365AcceptedDomain

        foreach ($dl in (Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)) {
            $dlName = Get-SafeProperty $dl 'DisplayName'
            $dlAddr = Get-SafeProperty $dl 'PrimarySmtpAddress'

            $members = @()
            try {
                $members = Get-DistributionGroupMember -Identity $dlAddr -ResultSize Unlimited -ErrorAction Stop
            } catch {
                Write-M365Log -Level Warning -Source 'exchange.distribution-groups' -Message "Members of $dlAddr : $($_.Exception.Message)"
            }

            $externals = @($members | Where-Object {
                Test-ExternalDomain -Address (Get-SafeProperty $_ 'PrimarySmtpAddress') -AcceptedDomains $accepted
            })

            if ($P.OnlyWithExternalMembers -and $externals.Count -eq 0) { continue }

            if (-not $P.IncludeMembers) {
                [pscustomobject]@{
                    GroupName          = $dlName
                    PrimarySmtpAddress = $dlAddr
                    GroupType          = Get-SafeProperty $dl 'GroupType'
                    MemberCount        = $members.Count
                    ExternalMemberCount= $externals.Count
                    ManagedBy          = (@(Get-SafeProperty $dl 'ManagedBy' @()) -join '; ')
                    HiddenFromGAL      = Get-SafeProperty $dl 'HiddenFromAddressListsEnabled'
                    RequireSenderAuth  = Get-SafeProperty $dl 'RequireSenderAuthenticationEnabled'
                    WhenCreated        = Get-SafeProperty $dl 'WhenCreated'
                }
                continue
            }

            foreach ($m in $members) {
                $addr = Get-SafeProperty $m 'PrimarySmtpAddress'
                [pscustomobject]@{
                    GroupName          = $dlName
                    PrimarySmtpAddress = $dlAddr
                    GroupType          = Get-SafeProperty $dl 'GroupType'
                    MemberCount        = $members.Count
                    MemberName         = Get-SafeProperty $m 'DisplayName'
                    MemberAddress      = $addr
                    MemberType         = Get-SafeProperty $m 'RecipientTypeDetails'
                    IsExternal         = Test-ExternalDomain -Address $addr -AcceptedDomains $accepted
                    ManagedBy          = (@(Get-SafeProperty $dl 'ManagedBy' @()) -join '; ')
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'exchange.dynamic-dl-members'
    Name     = 'Dynamic Distribution Group Members'
    Category = 'Exchange'
    Synopsis = 'Resolves the current membership of each dynamic distribution group.'
    Service  = 'ExchangeOnline'
    Risk     = 'ReadOnly'
    Replaces = 'Office 365 Dynamic Distribution Group Members Report (retired MSOnline dependency)'
    Notes    = 'Dynamic DL membership is evaluated at send time, so it has to be resolved from the recipient filter.'
    Parameters = @(
        @{ Name = 'GroupIdentity'; Type = 'String'
           Help = 'Optional: a single dynamic group. Blank for all.' }
    )
    Execute = {
        param($P)

        $groups = if ($P.GroupIdentity) {
            @(Get-DynamicDistributionGroup -Identity $P.GroupIdentity -ErrorAction Stop)
        } else {
            Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop
        }

        foreach ($g in $groups) {
            $gName = Get-SafeProperty $g 'DisplayName'
            $gAddr = Get-SafeProperty $g 'PrimarySmtpAddress'

            try {
                # Membership is derived by evaluating the group's recipient filter.
                $members = Get-Recipient -RecipientPreviewFilter (Get-SafeProperty $g 'RecipientFilter') `
                                         -ResultSize Unlimited -ErrorAction Stop
            } catch {
                Write-M365Log -Level Warning -Source 'exchange.dynamic-dl-members' -Message "Resolving $gAddr : $($_.Exception.Message)"
                continue
            }

            foreach ($m in $members) {
                [pscustomobject]@{
                    GroupName          = $gName
                    PrimarySmtpAddress = $gAddr
                    MemberCount        = @($members).Count
                    MemberName         = Get-SafeProperty $m 'DisplayName'
                    MemberAddress      = Get-SafeProperty $m 'PrimarySmtpAddress'
                    MemberType         = Get-SafeProperty $m 'RecipientTypeDetails'
                    RecipientFilter    = Get-SafeProperty $g 'RecipientFilter'
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'exchange.calendar-permissions'
    Name     = 'Calendar Permissions'
    Category = 'Exchange'
    Synopsis = 'Who can see or edit each mailbox calendar, including Default and Anonymous.'
    Service  = 'ExchangeOnline'
    Risk     = 'ReadOnly'
    Replaces = 'Get Calendar Permission Report'
    Parameters = @(
        @{ Name = 'MailboxFilter'; Type = 'String'
           Help = 'Optional: a single mailbox. Blank for all user mailboxes.' }
        @{ Name = 'OnlyNonDefault'; Type = 'Bool'; Default = $false
           Help = 'Hide the built-in Default/Anonymous entries at their standard values.' }
    )
    Execute = {
        param($P)

        $params = @{ ResultSize = 'Unlimited'; RecipientTypeDetails = 'UserMailbox'; ErrorAction = 'Stop' }
        if ($P.MailboxFilter) { $params.Remove('ResultSize'); $params.Remove('RecipientTypeDetails'); $params['Identity'] = $P.MailboxFilter }

        foreach ($mbx in (Get-Mailbox @params)) {
            $upn = Get-SafeProperty $mbx 'UserPrincipalName'

            # The calendar folder name is localised, so locate it by folder type.
            $calendarPath = "${upn}:\Calendar"
            try {
                $folder = Get-MailboxFolderStatistics -Identity $upn -FolderScope Calendar -ErrorAction Stop |
                          Where-Object { (Get-SafeProperty $_ 'FolderType') -eq 'Calendar' } | Select-Object -First 1
                if ($folder) {
                    $calendarPath = "${upn}:" + (Get-SafeProperty $folder 'FolderPath').Replace('/', '\')
                }
            } catch { }

            try {
                $perms = Get-MailboxFolderPermission -Identity $calendarPath -ErrorAction Stop
            } catch {
                Write-M365Log -Level Warning -Source 'exchange.calendar-permissions' -Message "Calendar of $upn : $($_.Exception.Message)"
                continue
            }

            foreach ($perm in $perms) {
                $user   = (Get-SafeProperty $perm 'User').ToString()
                $rights = (@(Get-SafeProperty $perm 'AccessRights' @()) -join ', ')

                if ($P.OnlyNonDefault -and $user -in @('Default','Anonymous') -and $rights -in @('None','AvailabilityOnly')) { continue }

                [pscustomobject]@{
                    Mailbox        = Get-SafeProperty $mbx 'DisplayName'
                    MailboxAddress = $upn
                    GrantedTo      = $user
                    AccessRights   = $rights
                    SharingFlags   = Get-SafeProperty $perm 'SharingPermissionFlags'
                    IsWideOpen     = [bool]($user -in @('Default','Anonymous') -and $rights -notin @('None','AvailabilityOnly'))
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
Register-M365Task @{
    Id       = 'exchange.mailbox-audit-config'
    Name     = 'Mailbox Audit Configuration'
    Category = 'Exchange'
    Synopsis = 'Which mailboxes have auditing on, and which actions are not being logged.'
    Service  = 'ExchangeOnline'
    Risk     = 'ReadOnly'
    Replaces = 'Enable Mailbox Auditing / Export Non-audited Mailbox Actions'
    Notes    = 'Gaps here become gaps in any later incident investigation.'
    Parameters = @(
        @{ Name = 'OnlyGaps'; Type = 'Bool'; Default = $true
           Help = 'Only mailboxes with auditing disabled or missing recommended actions.' }
    )
    Execute = {
        param($P)

        # Actions worth having on for investigations.
        $recommendedOwner    = @('HardDelete','SoftDelete','MoveToDeletedItems','UpdateInboxRules','MailboxLogin')
        $recommendedDelegate = @('HardDelete','SoftDelete','SendAs','SendOnBehalf','UpdateInboxRules')

        foreach ($mbx in (Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)) {
            $auditEnabled = Get-SafeProperty $mbx 'AuditEnabled' $false
            $owner    = @(Get-SafeProperty $mbx 'AuditOwner' @())
            $delegate = @(Get-SafeProperty $mbx 'AuditDelegate' @())

            $missingOwner    = @($recommendedOwner    | Where-Object { $_ -notin $owner })
            $missingDelegate = @($recommendedDelegate | Where-Object { $_ -notin $delegate })

            $hasGap = (-not $auditEnabled) -or $missingOwner.Count -gt 0 -or $missingDelegate.Count -gt 0
            if ($P.OnlyGaps -and -not $hasGap) { continue }

            [pscustomobject]@{
                DisplayName       = Get-SafeProperty $mbx 'DisplayName'
                UserPrincipalName = Get-SafeProperty $mbx 'UserPrincipalName'
                MailboxType       = Get-SafeProperty $mbx 'RecipientTypeDetails'
                AuditEnabled      = $auditEnabled
                AuditLogAgeDays   = Get-SafeProperty $mbx 'AuditLogAgeLimit'
                MissingOwnerActions    = ($missingOwner -join '; ')
                MissingDelegateActions = ($missingDelegate -join '; ')
                HasAuditGap       = $hasGap
            }
        }
    }
}
