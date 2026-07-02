# PBI Usage Export

A PowerShell script that extracts **Power BI / Microsoft Fabric activity events** (the tenant audit log) and exports them to an Excel workbook (`PBIUsage.xlsx`).

## ⚠️ Disclaimer

> **This is a test build created with GitHub Copilot.**
> It is provided **as-is**, with no warranty, and has not been hardened for production use.
> Review the code before running it, and use it carefully — especially since it reads
> **tenant-wide audit data** and requires **administrator** privileges. You are
> responsible for complying with your organization's data-handling and security policies.

## Goal

Give administrators a simple, repeatable way to pull Power BI / Fabric usage and audit
activity into a spreadsheet for review, reporting, or archiving — without needing the
Admin portal or a full monitoring solution.

## What it does

1. **Authenticates** interactively (system browser), by device code, or via Azure CLI.
2. **Calls the Power BI Admin "Get Activity Events" REST API** one UTC day at a time
   (the API only accepts a start/end within the same UTC day).
3. **Follows pagination** (`continuationUri`) until every page is retrieved.
4. **Flattens** each event (nested objects are JSON-stringified) into one row.
5. **Exports** everything to `PBIUsage.xlsx` in the script's folder as a multi-sheet,
   styled workbook (Overview dashboard, activity breakdown with a chart, per-activity
   user lists, and the full raw detail). See [Output workbook](#output-workbook) below.

## Output workbook

The script writes a single `.xlsx` with **five worksheets**, ordered from summary to
detail. Every sheet uses a coloured title banner, an Excel table (banded rows + filter
dropdowns), auto-sized columns, and frozen headers so they stay visible while scrolling.

| Sheet | Contents |
|---|---|
| **Overview** | A dashboard of key metrics: report generation time, date range (UTC), days collected, total events, unique users, distinct activities, and the counts of users who created and viewed reports. |
| **Activity Breakdown** | One row per `Activity` with its **event count** and **unique-user count**, sorted by frequency. Includes in-cell data bars and an *Events by Activity* column chart for a quick visual read. |
| **CreateReport Users** | The unique list of `UserId` (UPN) values that performed a `CreateReport` activity in the window. |
| **ViewReport Users** | The unique list of `UserId` (UPN) values that performed a `ViewReport` activity in the window. |
| **PBIUsage** | The full raw detail — **one row per audit event**, columns being the union of all fields seen across events (typically ~50+ columns such as `Id`, `CreationTime`, `Operation`, `Activity`, `UserId`, `WorkspaceId`, `WorkSpaceName`, `ObjectType`, `ObjectDisplayName`, `ResultStatus`, `CapacityName`, etc.). Fields not relevant to a given event are left blank. |

> If the target file is **open in Excel** (locked), the script writes to a timestamped
> copy (e.g. `PBIUsage_20260702_141530.xlsx`) instead of failing.
>
> If the `ImportExcel` module cannot be installed, the script falls back to writing a
> `.csv` (raw detail only) next to the requested output path.

## API used

**Power BI REST API – Admin – Get Activity Events**

```
GET https://api.powerbi.com/v1.0/myorg/admin/activityevents
      ?startDateTime='<UTC>'&endDateTime='<UTC>'
```

Docs: <https://learn.microsoft.com/rest/api/power-bi/admin/get-activity-events>

This endpoint returns the Microsoft 365 / Power BI **audit log**: one record per user or
system action across Power BI **and** Fabric (dataset refreshes, report views, notebook
runs, artifact create/update, Git commits, Copilot requests, and more). The schema is
sparse and operation-dependent — each record only carries the fields relevant to that
operation.

### Key constraints

| Constraint | Detail |
|---|---|
| **Permissions** | Requires a **Fabric / Power BI Tenant admin** identity (`Tenant.Read.All`). |
| **Time window** | `startDateTime` and `endDateTime` must fall within the **same UTC day**. |
| **Retention** | The service retains only the **last ~30 days**. Older days return HTTP 400. |

## Requirements

- **PowerShell** 5.1+ or PowerShell 7+.
- Modules (auto-installed for the current user if missing):
  - [`ImportExcel`](https://www.powershellgallery.com/packages/ImportExcel) — writes the `.xlsx`.
  - [`MSAL.PS`](https://www.powershellgallery.com/packages/MSAL.PS) — interactive / device-code sign-in.
- For `-AuthMode AzureCli`: the [Azure CLI](https://learn.microsoft.com/cli/azure/) (`az`).

> **Interactive sign-in** opens your **default system browser** (using an
> `http://localhost` loopback redirect), so the WebView2 runtime is **not** required.
> If the browser sign-in can't start, the script automatically falls back to device-code.

## Usage

```powershell
# Default: last 27 days -> PBIUsage.xlsx in the script folder, interactive browser sign-in
.\Get-PBIUsage.ps1

# Entire previous calendar month
.\Get-PBIUsage.ps1 -LastMonth

# Custom date range (UTC)
.\Get-PBIUsage.ps1 -StartDate 2026-06-15 -EndDate 2026-06-30

# Remote / headless session (device code sign-in)
.\Get-PBIUsage.ps1 -AuthMode DeviceCode

# Reuse an existing 'az login' session
.\Get-PBIUsage.ps1 -AuthMode AzureCli

# Custom output path and specific tenant
.\Get-PBIUsage.ps1 -Days 30 -OutputPath C:\temp\PBIUsage.xlsx -TenantId contoso.com
```

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-Days` | `27` | Days back from today (UTC). Ignored if `-LastMonth` or `-StartDate` is used. |
| `-LastMonth` | — | Collect the entire previous calendar month (UTC). |
| `-StartDate` / `-EndDate` | — | Explicit UTC date range (inclusive). |
| `-OutputPath` | `PBIUsage.xlsx` (script folder) | Target Excel file. |
| `-AuthMode` | `Interactive` | `Interactive` (system browser), `DeviceCode`, or `AzureCli`. |
| `-TenantId` | `organizations` | Tenant GUID or domain to sign in against. |
| `-AccessToken` | — | Bypass sign-in with a pre-acquired bearer token. |

## Expected outcome

- A sign-in prompt (unless you pass `-AccessToken` or reuse an `az` session).
- Console progress, one line per UTC day fetched, then a total event count.
- A styled Excel workbook at the output path with the five worksheets described in
  [Output workbook](#output-workbook) above.

Example console output:

```
Authenticated.
Collecting 27 day(s): 2026-06-05 to 2026-07-01 (UTC)
Fetching 2026-06-05 ...
...
Collected 1832 event(s).
Exported workbook: C:\temp\PBI-Usage\PBIUsage.xlsx
  Sheets: Overview | Activity Breakdown | CreateReport Users (1) | ViewReport Users (1) | PBIUsage (1832 rows)
```

> If the `ImportExcel` module cannot be installed, the script falls back to writing a
> `.csv` next to the requested output path.

## Notes & tips

- To reliably capture a **full calendar month**, run early in the following month (within
  the 30-day retention window) — e.g. schedule it for the 1st or 2nd.
- HTTP `400` on the oldest days is expected when they fall outside the retention window.
- `UserId = 00000009-0000-0000-c000-000000000000` indicates a **system/service** action
  (e.g. a scheduled refresh) rather than a person.

## License

MIT
