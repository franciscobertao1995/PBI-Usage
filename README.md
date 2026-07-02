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
5. **Exports** everything to `PBIUsage.xlsx` in the script's folder (worksheet `PBIUsage`,
   with autofilter, frozen header, and bold header row).

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
# Default: last 28 days -> PBIUsage.xlsx in the script folder, interactive browser sign-in
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
| `-Days` | `28` | Days back from today (UTC). Ignored if `-LastMonth` or `-StartDate` is used. |
| `-LastMonth` | — | Collect the entire previous calendar month (UTC). |
| `-StartDate` / `-EndDate` | — | Explicit UTC date range (inclusive). |
| `-OutputPath` | `PBIUsage.xlsx` (script folder) | Target Excel file. |
| `-AuthMode` | `Interactive` | `Interactive` (system browser), `DeviceCode`, or `AzureCli`. |
| `-TenantId` | `organizations` | Tenant GUID or domain to sign in against. |
| `-AccessToken` | — | Bypass sign-in with a pre-acquired bearer token. |

## Expected outcome

- A sign-in prompt (unless you pass `-AccessToken` or reuse an `az` session).
- Console progress, one line per UTC day fetched, then a total event count.
- An Excel file at the output path — **one row per audit event**, columns being the
  **union of all fields** seen across the events (typically ~50+ columns such as
  `Id`, `CreationTime`, `Operation`, `Activity`, `UserId`, `WorkspaceId`,
  `WorkSpaceName`, `ObjectType`, `ObjectDisplayName`, `ResultStatus`, `CapacityName`,
  etc.). Fields not relevant to a given event are left blank.

Example console output:

```
Authenticated.
Collecting 30 day(s): 2026-06-01 to 2026-06-30 (UTC)
Fetching 2026-06-01 ...
WARNING:   Day 2026-06-01 (page 0) failed: 400 (Bad Request).   # beyond 30-day retention
...
Collected 1635 event(s).
Exported 1635 rows to C:\temp\PBI-Usage\PBIUsage.xlsx
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
