# M365Toolkit

A task-driven Microsoft 365 administration toolkit with a WPF console, built to replace
the copy-pasted patterns in this repository's original script collection.

Everything is one module: one authentication layer, one task registry, one GUI. Adding a
capability means writing a task definition — no GUI changes required.

---

## Why this exists

The 168 standalone scripts in this repo each re-implement connection handling, each print
with `Write-Host`, and each write CSV from inside the script body. That makes them fine to
read and awkward to build on. See [`../REVIEW.md`](../REVIEW.md) for the full audit.

| Original pattern | Here |
|---|---|
| `Connect_MgGraph` helper copy-pasted into ~150 files | one `Connect-M365Toolkit` |
| Username/password "for scheduling" (basic auth — now disabled) | certificate-based auth only |
| `Write-Host` + `Export-Csv` inside the script | objects to the pipeline; caller exports |
| `Read-Host` prompts that hang scheduled runs | declared parameters with defaults |
| Retired `MSOnline` / `AzureAD` cmdlets | Microsoft Graph |
| No confirmation on destructive operations | risk levels + `WhatIf` by default |

---

## Requirements

- **PowerShell 5.1+** for the module; the GUI requires **Windows** (WPF).
- Modules, installed on demand: `Microsoft.Graph`, `ExchangeOnlineManagement`,
  `PnP.PowerShell`, `MicrosoftTeams`.

---

## Quick start

### GUI

```powershell
.\Start-M365Console.ps1
```

Connect, pick a task from the catalog, fill in the generated form, run. Results land in a
sortable grid with CSV/HTML export.

### Command line

```powershell
Import-Module .\M365Toolkit\M365Toolkit.psd1

# Interactive sign-in
Connect-M365Toolkit -Service Graph, ExchangeOnline

# What's available
Get-M365Task | Format-Table Id, Name, Category, Risk
Get-M365Task -Search 'forward'
Get-M365Task -Category Licensing

# Run something
Invoke-M365Task -Id 'licensing.unused-licenses' -Parameters @{ InactiveDays = 60 }
```

Because tasks emit objects, results compose:

```powershell
Invoke-M365Task -Id 'exchange.mailbox-sizes' |
    Where-Object UsagePercent -gt 85 |
    Sort-Object SizeGB -Descending |
    Export-M365Result -Path .\near-quota.html -Format Html -Title 'Mailboxes near quota'
```

### Unattended / scheduled

Certificate auth, no interactive prompt:

```powershell
Import-Module C:\M365Toolkit\M365Toolkit.psd1

Connect-M365Toolkit -Service Graph, ExchangeOnline `
    -TenantId     '00000000-0000-0000-0000-000000000000' `
    -ClientId     '11111111-1111-1111-1111-111111111111' `
    -CertificateThumbprint 'A1B2C3...' `
    -Organization 'contoso.onmicrosoft.com' `
    -LogFile      'C:\Reports\toolkit.log'

Invoke-M365Task -Id 'security.app-credential-expiry' -Parameters @{ DaysUntilExpiry = 30 } |
    Export-M365Result -Path "C:\Reports\app-creds-$(Get-Date -f yyyyMMdd).csv"

Disconnect-M365Toolkit
```

---

## Risk levels

Every task declares one. The GUI colour-codes them and confirms before running anything
that writes.

| Level | Meaning | Behaviour |
|---|---|---|
| `ReadOnly` | Reports only | Runs immediately |
| `Write` | Changes configuration | Confirmation prompt; `WhatIf` defaults to true |
| `Destructive` | Disables, removes, revokes | Confirmation prompt; `WhatIf` defaults to true |

Risky tasks take a `WhatIf` parameter that is **on by default** — the first run reports
what *would* change. Set it to false to apply.

```powershell
# Preview
Invoke-M365Task -Id 'security.block-external-forwarding' -Detailed

# Apply
Invoke-M365Task -Id 'security.block-external-forwarding' `
    -Parameters @{ WhatIf = $false } -Force
```

---

## Task catalog

33 tasks across six categories. `Get-M365Task` is the live list; each records which
original script it supersedes via its `Replaces` field.

### Security
`security.risky-users` · `security.failed-signins` · `security.external-forwarding` ·
`security.block-external-forwarding` · `security.compromised-remediation` ·
`security.app-credential-expiry` · `security.enterprise-app-owners`

### Identity
`identity.mfa-status` · `identity.mfa-methods-detail` · `identity.inactive-users` ·
`identity.guest-users` · `identity.admin-roles` · `identity.devices` ·
`identity.user-membership`

### Licensing
`licensing.sku-summary` · `licensing.unused-licenses` · `licensing.user-licenses` ·
`licensing.unlicensed-users` · `licensing.set-user-license`

### Exchange
`exchange.mailbox-sizes` · `exchange.mailbox-permissions` · `exchange.shared-mailboxes` ·
`exchange.distribution-groups` · `exchange.dynamic-dl-members` ·
`exchange.calendar-permissions` · `exchange.mailbox-audit-config`

### Lifecycle
`lifecycle.offboard-user` · `lifecycle.convert-to-shared` · `lifecycle.m365-groups`

### Sharing
`sharing.onedrive-usage` · `sharing.sharepoint-usage` · `sharing.anonymous-links` ·
`sharing.external-users`

---

## Adding a task

Drop a `Register-M365Task` call into any `*.tasks.ps1` file under `Tasks/`. It appears in
the GUI on next import, with its parameter form generated automatically.

```powershell
Register-M365Task @{
    Id       = 'exchange.big-attachments'
    Name     = 'Large Mailbox Items'
    Category = 'Exchange'
    Synopsis = 'Mailboxes holding unusually large single items.'
    Service  = 'ExchangeOnline'
    Risk     = 'ReadOnly'
    Parameters = @(
        @{ Name = 'MinimumMB'; Type = 'Int'; Default = 25; Help = 'Size floor in MB.' }
    )
    Execute = {
        param($P)
        foreach ($mbx in (Get-Mailbox -ResultSize Unlimited)) {
            $stats = Get-MailboxStatistics -Identity $mbx.UserPrincipalName
            [pscustomobject]@{
                Mailbox    = $mbx.DisplayName
                LargestMB  = [math]::Round((ConvertFrom-ExchangeSize $stats.TotalItemSize) / 1MB, 1)
            }
        }
    }
}
```

Parameter types: `String`, `Int`, `Bool`, `DateTime`, `Choice`, `MultiChoice`, `File`.
Definitions are validated at registration — a malformed task fails at import rather than
halfway through a tenant scan.

---

## Architecture notes

**Connection state is per-runspace.** `Connect-MgGraph` and `Connect-ExchangeOnline` store
their session in the runspace that called them. The GUI therefore keeps one long-lived
worker runspace and marshals every call — connection *and* task runs — through it, polling
completion on a `DispatcherTimer` so the window stays responsive during a tenant-wide scan.

**GUI helpers live at module scope**, not nested inside `Show-M365Console`. WPF handlers
created with `GetNewClosure()` capture variables but not function definitions, so a nested
helper can fail to resolve when a handler fires later.

**Tasks are stored as objects, not hashtables.** `Select-Object`, `Format-Table` and WPF
data binding all read properties and none of them read hashtable keys.

**No strict mode.** Graph and Exchange return objects whose property sets vary by tenant
licensing; strict mode turns a missing optional property into a hard failure mid-scan.
Tasks use `Get-SafeProperty` instead.

---

## Testing

```powershell
.\Tools\Test-M365Toolkit.ps1
```

25 tests covering registration, validation, parameter coercion, the risk gate, the
execution envelope, export and logging. No tenant connection required — connectivity is
stubbed. CI runs these plus PSScriptAnalyzer on every push.

---

## Command reference

| Command | Purpose |
|---|---|
| `Connect-M365Toolkit` | Connect one or more services |
| `Disconnect-M365Toolkit` | Disconnect and clear the auth context |
| `Get-M365Connection` | Live connection status per service |
| `Get-M365Task` | List/search the catalog |
| `Invoke-M365Task` | Run a task |
| `Register-M365Task` | Add a task to the registry |
| `Export-M365Result` | Export results as CSV, JSON or HTML |
| `Show-M365Console` | Open the GUI |
| `Get-M365Log` / `Clear-M365Log` | Read/clear the activity log |
