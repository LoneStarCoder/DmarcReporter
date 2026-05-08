# DMARC2

PowerShell tooling for pulling DMARC aggregate report attachments from Outlook, extracting XML payloads, normalizing them into JSON, enriching source IPs with GEO data, and generating CSV/HTML reporting artifacts.

## What This Repo Does

The codebase implements a local Windows pipeline:

1. Read DMARC report emails from an Outlook folder.
2. Save report attachments into `.\dmarc_report_emails`.
3. Extract `.zip`, `.gz`, and `.msg`-embedded archives into `.\dmarc_xml_exports`.
4. Parse DMARC XML into `mastertable.json`.
5. Optionally enrich source IPs into `GEOIP.json` and `mastertable.enriched.json`.
6. Generate CSV summaries and a simple HTML report in `.\DmarcReport`.
7. Generate a separate interactive dashboard in `.\DmarcDashboard\dashboard.html`.

## Requirements

- Windows PowerShell
- Outlook desktop installed on the machine
- An Outlook profile with access to the DMARC mailbox/folder
- Local filesystem write access to this directory
- Internet access if you want GEO lookups to run
- An IPInfo token if you want GEO enrichment

## Repository Layout

- `Invoke-DmarcReporter.ps1`
  Orchestrates email collection, archive extraction, XML processing, and optional GEO/report generation.

- `Get-Dmarc_emails.ps1`
  Connects to Outlook through COM, filters messages by date/read state, and saves attachments locally.

- `Extract-Dmarc_reports.ps1`
  Extracts attachments from `.msg` files and expands `.zip` / `.gz` DMARC report payloads into XML files.

- `Process-Dmarc_xml.ps1`
  Parses the XML exports and writes the normalized DMARC record set to `mastertable.json`.

- `Invoke-GEO_IP_Lookup.ps1`
  Looks up unique source IPs and writes/updates `GEOIP.json`.

- `Merge-GEOIntoMasterTable.ps1`
  Joins `GEOIP.json` back into `mastertable.json` and writes `mastertable.enriched.json`.

- `New-DMarcReport.ps1`
  Produces flat CSV outputs and a basic HTML report under `.\DmarcReport`.

- `Generate-Dashboard.ps1`
  Produces the richer standalone HTML dashboard under `.\DmarcDashboard`.

- `dmarc_report_emails\`
  Raw saved attachments from Outlook.

- `dmarc_xml_exports\`
  Extracted DMARC XML payloads.

- `DmarcReport\`
  Generated CSV and HTML reporting outputs.

- `DmarcDashboard\`
  Generated interactive dashboard output.

## Quick Start

Collect unread DMARC report attachments from a dedicated Outlook folder and generate the reporting set:

```powershell
.\Invoke-DmarcReporter.ps1 `
    -Days 7 `
    -EmailFolderPath 'Inbox\Ignore\dmarcreports' `
    -MessageFilter UnreadOnly `
    -MarkAsRead $true `
    -GenerateReports $true
```

Generate or refresh the dashboard separately:

```powershell
.\Generate-Dashboard.ps1
```

## Manual Pipeline

If you want to run each stage independently:

```powershell
.\Get-Dmarc_emails.ps1 -Days 7 -EmailFolderPath 'Inbox\Ignore\dmarcreports'
.\Extract-Dmarc_reports.ps1
.\Process-Dmarc_xml.ps1
.\Invoke-GEO_IP_Lookup.ps1 -ApiToken '<your token>'
.\Merge-GEOIntoMasterTable.ps1
.\New-DMarcReport.ps1
.\Generate-Dashboard.ps1
```

## Key Parameters

### `Invoke-DmarcReporter.ps1`

- `-Days`
  How many days of Outlook mail to inspect.

- `-EmailFolderPath`
  Outlook folder path. Current implementation expects a path starting with `Inbox\`.

- `-MessageFilter`
  `UnreadOnly` or `All`.

- `-MarkAsRead`
  Marks processed unread messages as read after attachment handling.

- `-GenerateReports`
  Runs GEO lookup, GEO merge, and `New-DMarcReport.ps1`.

### `Get-Dmarc_emails.ps1`

- `-DestinationFolder`
  Where saved attachments are written. Default: `.\dmarc_report_emails`

- `-ExistingFileAction`
  `Skip` or `Overwrite`

### `Invoke-GEO_IP_Lookup.ps1`

- `-InputJsonPath`
  Source record file. Default: `.\mastertable.json`

- `-OutputJsonPath`
  GEO cache file. Default: `.\GEOIP.json`

- `-ApiBaseUri`
  GEO provider base URI. Default is IPInfo.

- `-ApiToken`
  Token used for the GEO lookup provider.

- `-ForceRefresh`
  Re-query IPs already present in `GEOIP.json`.

### `New-DMarcReport.ps1`

- `-InputJsonPath`
  Prefers `.\mastertable.enriched.json` by default.

- `-OutputDirectory`
  Output folder for CSV/HTML artifacts. Default: `.\DmarcReport`

### `Generate-Dashboard.ps1`

- `-InputJsonPath`
  Defaults to `.\mastertable.enriched.json` when present, otherwise `.\mastertable.json`

- `-OutputPath`
  Default: `.\DmarcDashboard\dashboard.html`

- `-Title`
  Dashboard page title.

## Generated Outputs

### Core data

- `mastertable.json`
  Parsed DMARC records.

- `GEOIP.json`
  Cached GEO lookups by source IP.

- `mastertable.enriched.json`
  Parsed DMARC records with GEO fields merged in.

### `DmarcReport`

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

### `DmarcDashboard`

- `dashboard.html`

## Operational Notes

- `Generate-Dashboard.ps1` is not currently called by `Invoke-DmarcReporter.ps1`. Run it separately when you want the dashboard refreshed.
- `Get-Dmarc_emails.ps1` resolves Outlook folders by walking from `Inbox`, not from arbitrary mailbox roots.
- Generated files are written in place and are meant to be treated as build artifacts or working data.
- GEO enrichment is optional. Reporting still works against `mastertable.json` if GEO data is unavailable.

## Suggested Workflow

1. Route DMARC aggregate reports into a dedicated Outlook folder.
2. Run `Invoke-DmarcReporter.ps1` on a schedule or manually.
3. Run `Generate-Dashboard.ps1` after report generation if you need the dashboard refreshed.
4. Review `DmarcReport\*.csv` for exportable tabular data.
5. Open `DmarcReport\dmarc-report.html` or `DmarcDashboard\dashboard.html` for local review.

## Current Gaps

- No automated test suite is included.
- Outlook collection is Windows/desktop-Outlook specific because it depends on the COM object model.
- GEO lookups depend on an external service and local token management.
