# DMARC3 Implementation Plan

## Summary
Create `plan.md` first, then refactor DMARC3 from standalone scripts into a `DmarcDashboard` PowerShell module with a thin public runner, isolated per-run folders, config support, optional GEO enrichment, and tests. Use this file as the implementation checklist.

## Checklist
- [x] Create `plan.md` in the repo root containing this implementation plan and checklist.
- [x] Create the module structure: `DmarcDashboard.psd1`, `DmarcDashboard.psm1`, `Public\`, `Private\`, `config\`, and `tests\`.
- [x] Move reusable script logic into private functions while preserving current pipeline behavior.
- [x] Add `Invoke-DmarcDashboard` as the single all-in-one public command.
- [x] Add `New-DmarcDashboardConfig` and `config\config.example.json`.
- [x] Replace hardcoded relative outputs with isolated run folders under `Runs\<RunId>\`.
- [x] Remove the hardcoded IPInfo token and support config/env/SecretManagement token lookup.
- [x] Update README and `.gitignore` for module usage, config, generated runs, and local secrets.
- [x] Add Pester tests for pure helpers, XML conversion, GEO merge, config resolution, and local fixture-based pipeline pieces.
- [x] Run parse checks, Pester tests where available, and PSScriptAnalyzer if installed.

## Public Commands
- `Invoke-DmarcDashboard`
- `New-DmarcDashboardConfig`

## Private Pipeline Functions
- Initialize run folder and manifest.
- Resolve configuration.
- Collect Outlook attachments.
- Extract `.msg`, `.zip`, and `.gz` payloads.
- Convert DMARC XML to normalized JSON.
- Optionally perform GEO lookup.
- Optionally merge GEO data.
- Generate CSV/HTML report.
- Generate dashboard HTML.
- Finalize `run.json`.

## Production Layout Decision
- Root contains module metadata, documentation, config examples, tests, and planning notes.
- `Public\` contains exported commands only.
- `Private\` contains internal module helpers.
- `Private\PipelineScripts\` contains the former standalone pipeline scripts used as private implementation backends during this migration.
- New direct script entrypoints should not be added to root. Add public module commands instead.

## Run Folder Layout
- `input`
- `extracted`
- `normalized`
- `reports`
- `dashboard`
- `logs`
- `run.json`

## Testing
- Parse all `.ps1`, `.psm1`, and `.psd1` files.
- Add and run Pester tests for non-Outlook logic.
- Test local fixture processing without Outlook.
- Keep Outlook COM collection manually testable because it depends on desktop Outlook/profile state.
- If external GEO tests are needed, use mockable URI/token behavior and avoid requiring a live token by default.

## Assumptions
- `notes.md` is the intended replacement for missing `noted.md`.
- Current generated JSON/dashboard changes are user or runtime state and must not be reverted.
- GEO lookup should be optional and disabled unless configured.
- Existing behavior should remain available through the new module entrypoint.
