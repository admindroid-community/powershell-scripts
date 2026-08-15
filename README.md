## PowerShell Scripts for Microsoft 365 Management, Reporting, and Auditing

---

> ### 🛠️ This fork adds M365Toolkit — a GUI console for day-to-day M365 engineering
>
> The scripts below are a useful reference, but each one re-implements its own connection
> logic, prints with `Write-Host`, and writes CSV from inside the script body. Several
> depend on the **retired MSOnline / AzureAD modules** and no longer run at all.
>
> **[`M365Toolkit/`](M365Toolkit/README.md)** consolidates the most-used capabilities into
> one module with a WPF console: 33 tasks on modern Microsoft Graph and Exchange Online
> cmdlets, one shared certificate-based authentication layer, pipeline-first output, and
> confirmation gating on anything that writes.
>
> ```powershell
> .\Start-M365Console.ps1          # GUI
>
> Import-Module .\M365Toolkit\M365Toolkit.psd1
> Get-M365Task                     # browse the catalog
> Invoke-M365Task -Id 'licensing.unused-licenses'
> ```
>
> - **[REVIEW.md](REVIEW.md)** — full audit of all 168 scripts and what's broken
> - **[M365Toolkit/README.md](M365Toolkit/README.md)** — usage, task catalog, extending it
> - **[Tools/Invoke-RepoAudit.ps1](Tools/Invoke-RepoAudit.ps1)** — the repeatable audit
> - **[Tools/Test-M365Toolkit.ps1](Tools/Test-M365Toolkit.ps1)** — smoke tests

---

### Introduction

Welcome to our comprehensive PowerShell repository containing hundreds of scripts tailored for managing, reporting, and auditing Microsoft 365 environments. These scripts are designed to assist IT administrators in automating routine tasks, gathering detailed reports, and ensuring compliance across their Microsoft 365 tenant.

### Features

1. **Extensive Collection:** Over 100 PowerShell scripts for various tasks.
2. **Automation:** Automate mundane administrative tasks to save time and reduce errors.
3. **Reporting:** Generate detailed reports for auditing and compliance purposes.
4. **Management:** Simplify the management of Microsoft 365 resources and services.
5. **Auditing:** Monitor M365 activities to identify suspicious users and unauthorized accesses.
6. **Export:** Export the report into a well formatted CSV file.
7. **Scheduling:** Most scripts support scheduling capability to generate the report periodically.
8. **Customizable:** Easily modify scripts to fit specific needs.

### Script usage instruction

* Each script is self-contained and includes detailed comments and usage examples.
* Most scripts have built-in parameters and switch parameters to manage and report on your Microsoft 365 environment granularly.
* For more detailed use cases, refer to the linked blog within the script.

## Need more than what these scripts offer?? - **Try Free Microsoft 365 Administration Tool by AdminDroid**

AdminDroid's free Microsoft 365 management tool offers ***120+ essential reports, 60+ key management actions, and dashboards for free***. The report includes users, groups, group membership, licenses, license expiry, sign-in activities, password changes, license changes, MFA changes, admin role changes, etc.

Download [AdminDroid Microsoft 365 administration
tool](https://admindroid.com/download?src=GitHub) to experience the power of AdminDroid.

Additionally, AdminDroid Microsoft 365 administration tool provides ***3000+ pre-built reports, 100+ insightful dashboards, 450+ management actions*** and ***85+ ready to deploy alert policy templates*** to manage your Microsoft 365 organization effortlessly.



*Try out the demo to see the full range of features in action:* [*https://demo.admindroid.com*](https://demo.admindroid.com)

