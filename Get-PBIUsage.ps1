<#
.SYNOPSIS
    Extracts Power BI / Fabric activity events (audit log) via the Admin
    "Get Activity Events" REST API and exports them to an Excel file.

    API: https://learn.microsoft.com/rest/api/power-bi/admin/get-activity-events
    GET https://api.powerbi.com/v1.0/myorg/admin/activityevents

.DESCRIPTION
    * The API only accepts a start/end that fall within the SAME UTC day, so this
      script loops one day at a time across the requested range.
    * Results are paged via 'continuationUri' / 'continuationToken' -- the script
      follows every page until exhausted.
    * All events are flattened (nested objects are JSON-stringified) and written
      to PBIUsage.xlsx in the script's folder (one row per event).

.PARAMETER Days
    Number of days back from today (UTC) to collect. Default: 28.
    The service retains roughly the last 30 days. Ignored if -LastMonth or
    -StartDate/-EndDate are supplied.

.PARAMETER LastMonth
    Collect the entire previous calendar month (UTC). Overrides -Days.

.PARAMETER StartDate
    First day (inclusive, UTC) to collect. Use with -EndDate. Overrides -Days.

.PARAMETER EndDate
    Last day (inclusive, UTC) to collect. Defaults to today if only -StartDate
    is given.

.PARAMETER OutputPath
    Target Excel file. Default: PBIUsage.xlsx in the same folder as the script.

.PARAMETER AuthMode
    How to sign in each run:
      Interactive (default) - opens a browser sign-in prompt (MSAL.PS).
      DeviceCode            - shows a code to enter at https://microsoft.com/devicelogin
                              (useful on remote/headless sessions).
      AzureCli              - reuses an existing 'az login' session.
    You can still pass -AccessToken to bypass sign-in entirely.

.PARAMETER TenantId
    Optional tenant (GUID or domain) to sign in against. Default: 'organizations'.

.NOTES
    Requires:
      * An identity with Fabric/Power BI Tenant admin rights (Tenant.Read.All /
        Tenant.ReadWrite.All) OR that reads the admin monitoring workspace.
      * For Interactive/DeviceCode auth: the 'MSAL.PS' module (auto-installed for
        CurrentUser if missing). For AzureCli auth: the Azure CLI ('az').
      * The 'ImportExcel' PowerShell module (auto-installed for CurrentUser if
        missing). If installation is blocked, the script falls back to CSV.
#>

[CmdletBinding()]
param(
    [int]      $Days       = 28,
    [switch]   $LastMonth,
    [datetime] $StartDate,
    [datetime] $EndDate,
    [string]   $OutputPath = (Join-Path $PSScriptRoot 'PBIUsage.xlsx'),
    [ValidateSet('Interactive','DeviceCode','AzureCli')]
    [string] $AuthMode   = 'Interactive',
    [string] $TenantId   = 'organizations',
    [string] $AccessToken
)

$ErrorActionPreference = 'Stop'
$resource = 'https://analysis.windows.net/powerbi/api'
$apiBase  = 'https://api.powerbi.com/v1.0/myorg/admin/activityevents'
# Well-known Microsoft public client (Azure PowerShell) usable for interactive/device-code sign-in.
$clientId = '1950a258-227b-4e31-a9cf-717495945fc2'
$scopes   = @("$resource/.default")

function Get-PbiAccessToken {
    param([string] $Mode, [string] $Tenant, [string] $ClientId, [string[]] $Scopes, [string] $Resource)

    if ($Mode -eq 'AzureCli') {
        Write-Host 'Signing in via Azure CLI...' -ForegroundColor Cyan
        $t = az account get-access-token --resource $Resource --query accessToken -o tsv 2>$null
        if (-not $t) {
            az login --allow-no-subscriptions | Out-Null
            $t = az account get-access-token --resource $Resource --query accessToken -o tsv
        }
        if (-not $t) { throw "Azure CLI did not return a token. Try 'az login' manually." }
        return $t
    }

    # Interactive / DeviceCode -> MSAL.PS
    if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
        Write-Host 'Installing MSAL.PS module (CurrentUser)...' -ForegroundColor Cyan
        Install-Module MSAL.PS -Scope CurrentUser -Force -AllowClobber -AcceptLicense
    }
    Import-Module MSAL.PS

    if ($Mode -eq 'DeviceCode') {
        Write-Host 'Starting device-code sign-in...' -ForegroundColor Cyan
        $result = Get-MsalToken -ClientId $ClientId -TenantId $Tenant -Scopes $Scopes -DeviceCode
    } else {
        Write-Host 'Opening interactive browser sign-in...' -ForegroundColor Cyan
        # Use the default system browser (loopback redirect) instead of the embedded
        # WebView2 control, which requires the WebView2 runtime and fails on many hosts.
        try {
            $result = Get-MsalToken -ClientId $ClientId -TenantId $Tenant -Scopes $Scopes -Interactive -UseEmbeddedWebView:$false -RedirectUri 'http://localhost'
        } catch {
            Write-Warning "System-browser sign-in failed ($($_.Exception.Message)); falling back to device code."
            $result = Get-MsalToken -ClientId $ClientId -TenantId $Tenant -Scopes $Scopes -DeviceCode
        }
    }
    return $result.AccessToken
}

# --- 1. Acquire an access token (fresh sign-in each run) ------------------------
if (-not $AccessToken) {
    $AccessToken = Get-PbiAccessToken -Mode $AuthMode -Tenant $TenantId -ClientId $clientId -Scopes $scopes -Resource $resource
}
if (-not $AccessToken) { throw 'Failed to acquire an access token.' }
$headers = @{ Authorization = "Bearer $AccessToken" }
Write-Host 'Authenticated.' -ForegroundColor Green

# --- 2. Build the list of UTC days to query -------------------------------------
$today = (Get-Date).ToUniversalTime().Date
if ($LastMonth) {
    $firstOfThisMonth = Get-Date -Date $today -Day 1
    $rangeStart = $firstOfThisMonth.AddMonths(-1)
    $rangeEnd   = $firstOfThisMonth.AddDays(-1)
} elseif ($StartDate) {
    $rangeStart = $StartDate.Date
    $rangeEnd   = if ($EndDate) { $EndDate.Date } else { $today }
} else {
    $rangeStart = $today.AddDays(-$Days)
    $rangeEnd   = $today.AddDays(-1)
}

# Warn if any requested day is beyond the ~30-day retention window.
$retentionEdge = $today.AddDays(-30)
if ($rangeStart -lt $retentionEdge) {
    Write-Warning ("Requested start {0:yyyy-MM-dd} is older than the ~30-day retention window (edge {1:yyyy-MM-dd}). Older days may return no data." -f $rangeStart, $retentionEdge)
}

$dayList = @()
for ($dt = $rangeStart; $dt -le $rangeEnd; $dt = $dt.AddDays(1)) { $dayList += $dt }
Write-Host ("Collecting {0} day(s): {1:yyyy-MM-dd} to {2:yyyy-MM-dd} (UTC)" -f $dayList.Count, $rangeStart, $rangeEnd) -ForegroundColor Cyan

# --- 3. Loop day-by-day and page through results --------------------------------
$all = [System.Collections.Generic.List[object]]::new()

foreach ($dayDate in $dayList) {
    $day   = $dayDate.ToString('yyyy-MM-dd')
    $start = "'${day}T00:00:00.000Z'"
    $end   = "'${day}T23:59:59.999Z'"
    $uri   = "$apiBase`?startDateTime=$start&endDateTime=$end"

    Write-Host "Fetching $day ..." -ForegroundColor DarkGray
    $page = 0
    do {
        try {
            $resp = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        } catch {
            Write-Warning "  Day $day (page $page) failed: $($_.Exception.Message)"
            break
        }
        if ($resp.activityEventEntities) { $resp.activityEventEntities | ForEach-Object { $all.Add($_) } }
        $uri = $resp.continuationUri
        $page++
    } while ($uri)
}

Write-Host "Collected $($all.Count) event(s)." -ForegroundColor Green
if ($all.Count -eq 0) { Write-Warning 'No events found for the requested window. Nothing to export.'; return }

# --- 4. Flatten records (union of all fields; nested objects -> JSON) ------------
$columns = $all | ForEach-Object { $_.PSObject.Properties.Name } | Sort-Object -Unique
$flat = foreach ($e in $all) {
    $row = [ordered]@{}
    foreach ($c in $columns) {
        $v = $e.$c
        if ($null -ne $v -and ($v -is [System.Management.Automation.PSCustomObject] -or $v -is [System.Array])) {
            $v = ($v | ConvertTo-Json -Depth 10 -Compress)
        }
        $row[$c] = $v
    }
    [PSCustomObject]$row
}

# --- 5. Export to Excel (ImportExcel), fall back to CSV --------------------------
$outDir = Split-Path -Parent $OutputPath
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$haveImportExcel = Get-Module -ListAvailable -Name ImportExcel
if (-not $haveImportExcel) {
    Write-Host 'Installing ImportExcel module (CurrentUser)...' -ForegroundColor Cyan
    try {
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        $haveImportExcel = $true
    } catch {
        Write-Warning "Could not install ImportExcel: $($_.Exception.Message)"
        $haveImportExcel = $false
    }
}

if ($haveImportExcel) {
    Import-Module ImportExcel
    if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
    $flat | Export-Excel -Path $OutputPath -WorksheetName 'PBIUsage' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    Write-Host "Exported $($flat.Count) rows to $OutputPath" -ForegroundColor Green
} else {
    $csv = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
    $flat | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Warning "ImportExcel unavailable. Wrote CSV instead: $csv"
    Write-Host "(Open the CSV in Excel and 'Save As' .xlsx if you need the workbook format.)"
}
