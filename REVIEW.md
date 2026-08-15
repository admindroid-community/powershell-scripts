# Repository Review

A static audit of all **168 PowerShell scripts** across **165 folders** in this fork,
performed with the PowerShell 7.4 language parser (AST-based, not grep) plus targeted
runtime verification of suspect patterns.

Audit tooling lives in [`Tools/Invoke-RepoAudit.ps1`](Tools/Invoke-RepoAudit.ps1) so the
scan is repeatable.

---

## Summary

| Finding | Count | Severity |
|---|---:|---|
| Scripts that fail to parse | **0** | — |
| Scripts referencing retired AzureAD / MSOnline modules | **13** | Critical |
| Scripts whose Exchange Online auth path uses retired Basic Auth | **81** | Critical |
| Scripts handling a credential as plaintext (any service) | **103** | Warning |
| Scripts requiring interactive `Read-Host` (block scheduling) | **160** | High |
| Scripts emitting results only via `Write-Host` / direct `Export-Csv` | **100** | High |
| Scripts with no error handling at all | **45** | Medium |
| Scripts using `+=` array append inside a loop (O(n²)) | **52** | Medium |
| Scripts with placeholder values left in | **11** | Low |
| Scripts supporting certificate-based auth | 116 | (good) |

Service coverage: Exchange Online 87 · Microsoft Graph 61 · PnP 12 · Teams 7 · SPO 4 · Security & Compliance 1.

---

## Critical findings

### 1. Thirteen scripts depend on retired modules

Microsoft retired the **MSOnline** and **AzureAD** PowerShell modules. Scripts calling
`Connect-MsolService` or `Connect-AzureAD` no longer authenticate.

| Script | Retired dependency |
|---|---|
| `Connect to All Office 365 Services/ConnectO365Services.ps1` | AzureAD + MSOnline |
| `Enable MFA for Admin Users/EnableMFAforAdmins.ps1` | MSOnline |
| `Identify MFA Deployment Source/IdentifyMFADeploymentSourcesReport.ps1` | MSOnline |
| `LicenseExpiryDateReport/LicenseExpiryDateReport.ps1` | MSOnline |
| `Mail Traffic Report/MailTrafficReport.ps1` | MSOnline |
| `Microsoft 365 Group Report/M365GroupReport.ps1` | MSOnline |
| `Office 365 Dynamic Distribution Group Members Report/…` | MSOnline |
| `Office 365 Mailbox Permissions Report/Prerequisites.ps1` | MSOnline |
| `Office 365 User Last Activity Time Report/*.ps1` | MSOnline |
| `Office 365 User Last Logon Time Report/Prerequisites.ps1` | MSOnline |
| `Office 365 User MFA Status Report/GetMFAStatus.ps1` | MSOnline |
| `Office365 License Reporting And Management/…` | MSOnline |
| `Azure AD Devices Report/GetAzureADDevicesReport.ps1` | AzureAD |

**Replacement mapping** (implemented in the toolkit shipped with this review):

| Retired | Modern replacement |
|---|---|
| `Connect-MsolService` / `Connect-AzureAD` | `Connect-MgGraph` |
| `Get-MsolUser` / `Get-AzureADUser` | `Get-MgUser` |
| `Get-MsolAccountSku` | `Get-MgSubscribedSku` |
| `Set-MsolUserLicense` | `Set-MgUserLicense` |
| `Get-MsolUser -All | Select StrongAuthenticationMethods` | `Get-MgUserAuthenticationMethod` |
| `Get-AzureADDevice` | `Get-MgDevice` |
| `Get-AzureADGroupMember` | `Get-MgGroupMember` |
| `Get-AzureADServicePrincipal` | `Get-MgServicePrincipal` |
| `Get-AzureADApplication` | `Get-MgApplication` |
| `Get-MsolCompanyInformation` | `Get-MgOrganization` |

### 2. The "scheduling" auth path is dead

103 scripts build a credential from a plaintext password; in 81 of them that credential is
passed to `Connect-ExchangeOnline`. They share this pattern:

```powershell
$SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
$Credential = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
Connect-ExchangeOnline -Credential $Credential
```

The inline comment says *"Storing credential in script for scheduling purpose —
Authentication using non-MFA account."* Two problems:

1. **Basic authentication is disabled in Exchange Online.** This path fails outright.
2. It encourages storing a plaintext admin password in a script file.

Certificate-based auth is the supported replacement, and 116 scripts already accept
`-CertificateThumbprint`. The plaintext branch should simply be deleted.

### 3. `SharePointDLMapping.ps1` is non-functional

```powershell
$URL = “https://your_domain.sharepoint.com/...”
$IESession = Start-Process -file iexplore -arg $URL -PassThru -WindowStyle Hidden
```

Drives **Internet Explorer**, which is retired and removed from supported Windows builds,
then maps a drive via the `WScript.Network` COM object. It also carries an unedited
placeholder URL. This cannot work on a current machine.

> Note: the curly quotes here are *not* a defect — PowerShell accepts U+201C/U+201D as
> string delimiters. This was verified at runtime rather than assumed.

---

## Architectural findings

These are why the collection is hard to build on, as opposed to individually buggy.

### Every script re-implements connection logic
There is no shared module. `Connect_MgGraph` / `Connect_EXO` helper functions are
copy-pasted across ~150 files, each with its own module-install prompt. A fix to one is a
fix to one.

### Results are terminal output, not objects
100 scripts print with `Write-Host` and write results straight to `Export-Csv` inside the
script body. Nothing is emitted to the pipeline, so results cannot be filtered, joined,
piped into another report, or consumed by a GUI without re-parsing a CSV off disk.

```powershell
# current: terminal-only, uncomposable
$Results | Select-Object 'Display Name',… | Export-Csv -Path $ExportCSV -Notype -Append
```

The fix is to emit objects and let the *caller* decide on formatting or export.

### Interactive prompts block automation
160 scripts call `Read-Host`, often for required parameters. Combined with the "scheduler
friendly" claim in the headers, this is contradictory — a scheduled run hangs on the prompt.

### O(n²) accumulation
52 scripts append with `+=` inside loops. On a 50,000-mailbox tenant this dominates runtime.
`List<T>` or pipeline output is O(n).

### No repo hygiene
No `.gitignore` (scripts drop CSVs into the working directory), no CI, no linting, no tests,
no manifest, and no index — finding the right script among 165 similarly-named folders is
manual.

---

## What was done about it

Rather than patch 168 files individually, this review ships a consolidated toolkit that
implements the corrected patterns:

- **`M365Toolkit/`** — a proper PowerShell module: one shared authentication layer
  (Graph / EXO / PnP / Teams) with certificate-based auth, a task registry, pipeline-first
  output, structured logging, and a WPF GUI console.
- **`Tools/Invoke-RepoAudit.ps1`** — the repeatable audit that produced this report.
- **`.gitignore`** — keeps generated CSV/HTML reports out of the repo.
- **`.github/workflows/lint.yml`** — PSScriptAnalyzer on every push.

See [`M365Toolkit/README.md`](M365Toolkit/README.md) for usage.

The original scripts are left untouched — they remain a useful reference, and the
toolkit's task definitions record which script each one supersedes via a `Replaces` field.
