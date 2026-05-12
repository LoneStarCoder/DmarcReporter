I want this to be an easy to use all in one solution.
I think it would be best if each run is a complete run on its own.
Each run would use a unique folder, so as to not mess up history.

Running without Variables would be good, except I think that it will always have to have the one variable for the Folder, since we are using outlook.

A single run would do everything.

Hardcoding needs to be changed into .env and configuration files.

How do we accomplish this?
Do we create a DMARC module?
Do we combine everything into a single script?
Do we use a master script to call everything?

------------------------------------------------

Build it as **both**:

1. **A PowerShell module** for reusable logic.
2. **One master runner script** for easy public use.

Do **not** combine everything into one giant script. That becomes hard to test, maintain, and publish.

## Recommended public design

```text
DmarcDashboard/
  DmarcDashboard.psd1
  DmarcDashboard.psm1
  Public/
    Invoke-DmarcDashboard.ps1
    New-DmarcDashboardConfig.ps1
  Private/
    Get-OutlookDmarcAttachment.ps1
    Convert-DmarcXmlReport.ps1
    Export-DmarcData.ps1
    New-DmarcHtmlDashboard.ps1
    Write-DmarcRunLog.ps1
  config/
    config.example.json
  .env.example
  README.md
  LICENSE
  tests/
```

PowerShell Gallery publishing requires a module manifest `.psd1`, and Microsoft recommends manifests for versioning, dependencies, compatibility, and publishing metadata. ([Microsoft Learn][1])

## How users should run it

Primary easy mode:

```powershell
Invoke-DmarcDashboard -MailboxFolder "DMARC"
```

Better repeatable mode:

```powershell
Invoke-DmarcDashboard -ConfigPath .\config\dmarc.config.json
```

Advanced mode:

```powershell
Invoke-DmarcDashboard `
  -MailboxFolder "DMARC" `
  -OutputRoot "C:\DMARC\Runs" `
  -StartDate "2026-05-01" `
  -EndDate "2026-05-12"
```

## Run model

Your instinct is right: **each run should be isolated**.

Example:

```text
Runs/
  2026-05-12_143012_8f3a/
    input/
    extracted/
    normalized/
    reports/
    dashboard/
    logs/
    run.json
```

Each run should produce a `run.json` manifest containing:

```json
{
  "RunId": "2026-05-12_143012_8f3a",
  "StartedAt": "2026-05-12T14:30:12Z",
  "ConfigUsed": {},
  "MailboxFolder": "DMARC",
  "OutputPaths": {},
  "ReportCount": 0,
  "Errors": []
}
```

This gives history without cross-run contamination.

## Configuration decision

Use **JSON config for normal settings** and **PowerShell SecretManagement or environment variables for secrets**.

Do not put secrets in `.env` as the main recommendation. `.env` is acceptable for local/dev convenience, but for a public security-adjacent tool, recommend SecretManagement. Microsoft’s SecretManagement module is designed to retrieve secrets from registered vaults. ([Microsoft Learn][2])

Example `config.example.json`:

```json
{
  "MailboxFolder": "DMARC",
  "OutputRoot": "C:\\DMARC\\Runs",
  "KeepRawAttachments": true,
  "CreateHtmlDashboard": true,
  "CreateCsv": true,
  "CreateJson": true,
  "TimeZone": "UTC"
}
```

## Recommended architecture

Use this as the only public command:

```powershell
Invoke-DmarcDashboard
```

Internally it calls:

```powershell
Initialize-DmarcRun
Get-DmarcMessages
Save-DmarcAttachments
Expand-DmarcArchives
Convert-DmarcAggregateReport
Export-DmarcDataset
New-DmarcDashboard
Complete-DmarcRun
```

Public users see one simple command. Maintainers get clean internals.

## Dependencies

Minimize dependencies, but document them clearly.

Likely dependencies:

```powershell
#Requires -Version 7.2
#Requires -Modules Microsoft.PowerShell.SecretManagement
```

Use `#Requires` for required PowerShell version/modules where appropriate. Microsoft documents `#Requires` as a way to require modules, versions, assemblies, and elevated privileges before execution. ([Microsoft Learn][3])

Use PSScriptAnalyzer and Pester before publishing. PSScriptAnalyzer checks PowerShell code quality against best-practice rules, and Pester is the standard PowerShell test framework. ([Microsoft Learn][4])

## Decision answers

| Question                                 | Recommendation                                                         |
| ---------------------------------------- | ---------------------------------------------------------------------- |
| Create a DMARC module?                   | **Yes.** Best for public distribution, testing, versioning, and reuse. |
| Combine everything into a single script? | **No.** Only acceptable for a quick prototype.                         |
| Use a master script?                     | **Yes.** But make it a thin entrypoint: `Invoke-DmarcDashboard`.       |
| Use `.env`?                              | **Only optional.** Prefer JSON config + SecretManagement.              |
| One complete run per execution?          | **Yes.** Isolated run folders are the right design.                    |
| Require folder variable?                 | **Yes, unless a config file supplies it.**                             |

## My recommended path

Publish it as a PowerShell module named something like:

```text
DmarcDashboard
```

Expose only a few public commands:

```powershell
Invoke-DmarcDashboard
New-DmarcDashboardConfig
Get-DmarcDashboardRun
```

Everything else should be private helper functions.

This gives you the best balance: easy for beginners, clean for advanced users, and maintainable enough to publish publicly.

[1]: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_module_manifests?view=powershell-7.6&utm_source=chatgpt.com "about_Module_Manifests - PowerShell | Microsoft Learn"
[2]: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.secretmanagement/?view=ps-modules&utm_source=chatgpt.com "Microsoft.PowerShell.SecretManagement Module - PowerShell"
[3]: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_requires?view=powershell-7.6&utm_source=chatgpt.com "about_Requires - PowerShell | Microsoft Learn"
[4]: https://learn.microsoft.com/en-us/powershell/utility-modules/psscriptanalyzer/overview?view=ps-modules&utm_source=chatgpt.com "PSScriptAnalyzer module - PowerShell | Microsoft Learn"
