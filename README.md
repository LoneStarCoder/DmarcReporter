# DMARCReporter

PowerShell tooling for pulling DMARC aggregate report attachments from Outlook, extracting XML payloads, normalizing them into JSON, enriching source IPs with GEO data, and generating CSV/HTML reporting artifacts.

- Author(s): Brody Kilpatrick (LonestarCoder), Codex
- Github Repository: https://github.com/LoneStarCoder
- License: https://github.com/LoneStarCoder/DmarcReporter/blob/main/LICENSE

## Attribution

Some DMARC processing concepts and modified helper logic were inspired by Martijn van Geffen's MIT-licensed [`DMARC-report`](https://github.com/martijnvangeffen/DMARC-report) PowerShell project.

## Module Quick Start

Import the module from the repo root:

```powershell
Import-Module .\DMARCReporter.psd1 -Force
```

Create a starter config:

```powershell
New-DMARCReporterConfig -Path .\config\dmarc.config.json -MailboxFolder 'Inbox\dmarcreports'
```

Run the full pipeline:

```powershell
Invoke-DMARCReporter -ConfigPath .\config\dmarc.config.json
```

Or run without a config file:

```powershell
Invoke-DMARCReporter `
    -MailboxFolder 'Inbox\dmarcreports' `
    -OutputRoot '.\Runs' `
    -Days 7 `
    -MessageFilter UnreadOnly
```

Each module run writes to an isolated folder:

```text
Runs\<RunId>\
  input\
  extracted\
  normalized\
  reports\
  dashboard\
  logs\
  run.json
```

GEO lookup is optional and disabled by default, but the runner will use an existing local `.\GEOIP.json` cache by default when present. To refresh or add missing cache entries, provide a token through config, the `DMARC_REPORTER_IPINFO_TOKEN` environment variable, or a SecretManagement secret named `DMARCReporter-IpInfoToken`, then run with `-EnableGeoLookup`. The local `GEOIP.json` cache is ignored by git.

## Setting The IPInfo API Token

The recommended option is to store the token in an environment variable named `DMARC_REPORTER_IPINFO_TOKEN`:

```powershell
[Environment]::SetEnvironmentVariable('DMARC_REPORTER_IPINFO_TOKEN','<your-ipinfo-token>','User')
```

Open a new PowerShell session after setting the user environment variable, then run:

```powershell
Import-Module .\DMARCReporter.psd1 -Force
Invoke-DMARCReporter -ConfigPath .\config\dmarc.config.json -EnableGeoLookup
```

For a current-session-only token, use:

```powershell
$env:DMARC_REPORTER_IPINFO_TOKEN = '<your-ipinfo-token>'
Invoke-DMARCReporter -ConfigPath .\config\dmarc.config.json -EnableGeoLookup
```

You can also use PowerShell SecretManagement:

```powershell
Set-Secret -Name 'DMARCReporter-IpInfoToken' -Secret '<your-ipinfo-token>'
Invoke-DMARCReporter -ConfigPath .\config\dmarc.config.json -EnableGeoLookup
```

For a one-off run, pass the token directly:

```powershell
Invoke-DMARCReporter `
    -ConfigPath .\config\dmarc.config.json `
    -EnableGeoLookup `
    -GeoApiToken '<your-ipinfo-token>'
```

Avoid committing a real token to `config\dmarc.config.json`. If you still want to configure token lookup there, set only the lookup names:

```json
{
  "EnableGeoLookup": true,
  "GeoApiTokenEnvName": "DMARC_REPORTER_IPINFO_TOKEN",
  "GeoApiTokenSecretName": "DMARCReporter-IpInfoToken"
}
```

## What This Repo Does

The codebase implements a local Windows pipeline:

1. Read DMARC report emails from an Outlook folder.
2. Save report attachments into the run `input\` folder.
3. Extract `.zip`, `.gz`, and `.msg`-embedded archives into the run `extracted\` folder.
4. Parse DMARC XML into `normalized\mastertable.json`.
5. Optionally enrich source IPs into `normalized\GEOIP.json` and `normalized\mastertable.enriched.json`.
6. Generate CSV summaries and a simple HTML report in the run `reports\` folder.
7. Generate a separate interactive dashboard in the run `dashboard\dashboard.html`.

## Requirements

- Windows PowerShell
- Outlook desktop installed on the machine
- An Outlook profile with access to the DMARC mailbox/folder
- Local filesystem write access to this directory
- Internet access if you want GEO lookups to run
- An IPInfo token if you want GEO enrichment

## Repository Layout

- `DMARCReporter.psd1` / `DMARCReporter.psm1`
  Module manifest and loader.

- `Public\Invoke-DMARCReporter.ps1`
  Main all-in-one public runner.

- `Public\New-DMARCReporterConfig.ps1`
  Creates a starter config file.

- `Private\`
  Internal module helpers for config, run initialization, token lookup, and pipeline orchestration.

- `Private\PipelineScripts\`
  Private implementation backends for Outlook collection, archive extraction, XML normalization, GEO enrichment, report generation, and dashboard rendering.

- `config\config.example.json`
  Example config. Local `config\*.json` files are ignored by git except the example.

- `tests\`
  Pester tests and local fixtures.

- `Runs\`
  Generated run output. Ignored by git.

## Key Parameters

### `Invoke-DMARCReporter`

- `-Days`
  How many days of Outlook mail to inspect.

- `-MailboxFolder`
  Outlook folder path. Current implementation expects a path starting with `Inbox\`.

- `-MessageFilter`
  `UnreadOnly` or `All`.

- `-MarkAsRead`
  Marks processed unread messages as read after attachment handling.

- `-ExistingFileAction`
  `Skip` or `Overwrite`

- `-EnableGeoLookup`
  Refreshes or creates the local GEO cache and enriches the run output.

- `-GeoApiToken`
  One-off IPInfo token value. Prefer environment variables or SecretManagement for regular use.

- `-SkipReports` / `-SkipDashboard`
  Skips report or dashboard generation when only normalized data is needed.

## Generated Outputs

### Core data

- `normalized\mastertable.json`
  Parsed DMARC records.

- `normalized\GEOIP.json`
  Cached GEO lookups by source IP.

- `normalized\mastertable.enriched.json`
  Parsed DMARC records with GEO fields merged in.

### `reports`

- `dmarc-detail-normalized.csv`
- `dmarc-mastertable-raw.csv`
- `dmarc-mastertable-flattened.csv`
- `dmarc-quarantine-simulation.csv`
- `dmarc-summary-by-org.csv`
- `dmarc-summary-by-sourceip.csv`
- `dmarc-summary-by-auth-result.csv`
- `dmarc-summary-by-spf-domain.csv`
- `dmarc-summary-by-dkim-domain.csv`
- `dmarc-auth-failures.csv`
- `dmarc-report.html`

### `dashboard`

- `dashboard.html`

## Operational Notes

- `Invoke-DMARCReporter` generates reports and dashboard output in one run.
- `Get-Dmarc_emails.ps1` resolves Outlook folders by walking from `Inbox`, not from arbitrary mailbox roots.
- Generated files are written under `Runs\` and are ignored by git.
- GEO enrichment is optional. Reporting still works against `mastertable.json` if GEO data is unavailable.

## Suggested Workflow

1. Route DMARC aggregate reports into a dedicated Outlook folder.
2. Run `Invoke-DMARCReporter` on a schedule or manually.
3. Review `Runs\<RunId>\reports\*.csv` for exportable tabular data.
4. Open `Runs\<RunId>\reports\dmarc-report.html` or `Runs\<RunId>\dashboard\dashboard.html` for local review.

## Current Gaps

- Outlook collection is Windows/desktop-Outlook specific because it depends on the COM object model.
- GEO lookups depend on an external service and local token management.

## License

This project is licensed under the MIT License. See `LICENSE`.


