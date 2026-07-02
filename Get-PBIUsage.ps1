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
    Number of days back from today (UTC) to collect. Default: 27.
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
    [int]      $Days       = 27,
    [switch]   $LastMonth,
    [datetime] $StartDate,
    [datetime] $EndDate,
    [string]   $OutputPath,
    [ValidateSet('Interactive','DeviceCode','AzureCli')]
    [string] $AuthMode   = 'Interactive',
    [string] $TenantId   = 'organizations',
    [string] $AccessToken
)

$ErrorActionPreference = 'Stop'

# Resolve the default output path. $PSScriptRoot can be empty depending on how the
# script is invoked, so fall back to the script file's folder, then the current dir.
if (-not $OutputPath) {
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir -and $PSCommandPath) { $scriptDir = Split-Path -Parent $PSCommandPath }
    if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
    $OutputPath = Join-Path $scriptDir 'PBIUsage.xlsx'
}

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
        $installArgs = @{ Name = 'MSAL.PS'; Scope = 'CurrentUser'; Force = $true; AllowClobber = $true }
        # -AcceptLicense only exists in PowerShellGet 1.6.0+ (Windows PowerShell 5.1
        # ships with 1.0.0.1, which lacks it). Add it only when supported.
        if ((Get-Command Install-Module).Parameters.ContainsKey('AcceptLicense')) {
            $installArgs['AcceptLicense'] = $true
        }
        Install-Module @installArgs
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

    # If the target workbook is open/locked (e.g. in Excel), remove it; on
    # failure fall back to a timestamped file so the run still succeeds.
    if (Test-Path $OutputPath) {
        try {
            Remove-Item $OutputPath -Force -ErrorAction Stop
        } catch {
            $alt = [System.IO.Path]::Combine(
                (Split-Path -Parent $OutputPath),
                ('{0}_{1}{2}' -f [System.IO.Path]::GetFileNameWithoutExtension($OutputPath),
                    (Get-Date -Format 'yyyyMMdd_HHmmss'),
                    [System.IO.Path]::GetExtension($OutputPath)))
            Write-Warning "'$OutputPath' is locked (open in Excel?). Writing to '$alt' instead."
            $OutputPath = $alt
        }
    }

    $titleColor = [System.Drawing.Color]::FromArgb(31, 78, 120)   # dark blue
    $barColor   = [System.Drawing.Color]::FromArgb(0, 112, 192)   # medium blue

    # Merge + colour the title cell into a full-width banner.
    function Set-TitleBar {
        param($Worksheet, [int] $LastCol)
        $cells = $Worksheet.Cells[1, 1, 1, $LastCol]
        $cells.Merge = $true
        $cells.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cells.Style.Fill.BackgroundColor.SetColor($titleColor)
        $cells.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $cells.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $Worksheet.Row(1).Height = 24
        # Freeze the title (row 1) and header (row 2) so the table header stays visible.
        $Worksheet.View.FreezePanes(3, 1)
    }

    # --- Aggregations ------------------------------------------------------
    $getUsersFor = {
        param($ActivityName)
        @($flat |
            Where-Object { $_.Activity -eq $ActivityName -and $_.UserId } |
            Select-Object -ExpandProperty UserId |
            Sort-Object -Unique)
    }
    $createReportUsers = & $getUsersFor 'CreateReport'
    $viewReportUsers   = & $getUsersFor 'ViewReport'

    $activityBreakdown = $flat |
        Group-Object Activity |
        Sort-Object Count -Descending |
        ForEach-Object {
            [PSCustomObject]@{
                Activity       = $_.Name
                'Event Count'  = $_.Count
                'Unique Users' = @($_.Group | Where-Object UserId |
                                    Select-Object -ExpandProperty UserId -Unique).Count
            }
        }

    $uniqueUsers = @($flat | Where-Object UserId |
                        Select-Object -ExpandProperty UserId -Unique).Count

    $overview = @(
        [PSCustomObject]@{ Metric = 'Report generated (local time)'; Value = (Get-Date).ToString('yyyy-MM-dd HH:mm') }
        [PSCustomObject]@{ Metric = 'Date range (UTC)';              Value = ('{0:yyyy-MM-dd}  to  {1:yyyy-MM-dd}' -f $rangeStart, $rangeEnd) }
        [PSCustomObject]@{ Metric = 'Days collected';               Value = $dayList.Count }
        [PSCustomObject]@{ Metric = 'Total events';                 Value = $flat.Count }
        [PSCustomObject]@{ Metric = 'Unique users';                 Value = $uniqueUsers }
        [PSCustomObject]@{ Metric = 'Distinct activities';          Value = $activityBreakdown.Count }
        [PSCustomObject]@{ Metric = 'Users who created reports (CreateReport)'; Value = $createReportUsers.Count }
        [PSCustomObject]@{ Metric = 'Users who viewed reports (ViewReport)';    Value = $viewReportUsers.Count }
    )

    # --- 1) Overview -------------------------------------------------------
    $pkg = $overview | Export-Excel -Path $OutputPath -WorksheetName 'Overview' -PassThru `
        -Title 'Power BI Usage Report - Overview' -TitleBold -TitleSize 14 `
        -TableName 'tblOverview' -TableStyle Medium2 -AutoSize
    Set-TitleBar $pkg.Workbook.Worksheets['Overview'] 2

    # --- 2) Activity Breakdown (table + data bars + chart) -----------------
    $n = $activityBreakdown.Count
    $activityBreakdown | Export-Excel -ExcelPackage $pkg -WorksheetName 'Activity Breakdown' `
        -Title 'Activity Breakdown' -TitleBold -TitleSize 14 `
        -TableName 'tblActivity' -TableStyle Medium2 -AutoSize -PassThru | Out-Null
    $wsAct = $pkg.Workbook.Worksheets['Activity Breakdown']
    Set-TitleBar $wsAct 3
    Add-ConditionalFormatting -Worksheet $wsAct -Range "B3:B$(2 + $n)" -DataBarColor $barColor
    Add-ConditionalFormatting -Worksheet $wsAct -Range "C3:C$(2 + $n)" -DataBarColor $barColor
    try {
        Add-ExcelChart -Worksheet $wsAct -ChartType ColumnClustered -Title 'Events by Activity' `
            -XRange "'Activity Breakdown'!A3:A$(2 + $n)" -YRange "'Activity Breakdown'!B3:B$(2 + $n)" `
            -Width 640 -Height 340 -Row 1 -Column 4
    } catch {
        Write-Warning "Chart skipped: $($_.Exception.Message)"
    }

    # --- 3) CreateReport Users --------------------------------------------
    $crRows = if ($createReportUsers.Count) {
        $createReportUsers | ForEach-Object { [PSCustomObject]@{ 'User (UPN)' = $_ } }
    } else {
        , [PSCustomObject]@{ 'User (UPN)' = '(no CreateReport activity found)' }
    }
    $crRows | Export-Excel -ExcelPackage $pkg -WorksheetName 'CreateReport Users' `
        -Title "Users who created reports ($($createReportUsers.Count))" -TitleBold -TitleSize 14 `
        -TableName 'tblCreateReport' -TableStyle Medium2 -AutoSize -PassThru | Out-Null
    Set-TitleBar $pkg.Workbook.Worksheets['CreateReport Users'] 1

    # --- 4) ViewReport Users ----------------------------------------------
    $vrRows = if ($viewReportUsers.Count) {
        $viewReportUsers | ForEach-Object { [PSCustomObject]@{ 'User (UPN)' = $_ } }
    } else {
        , [PSCustomObject]@{ 'User (UPN)' = '(no ViewReport activity found)' }
    }
    $vrRows | Export-Excel -ExcelPackage $pkg -WorksheetName 'ViewReport Users' `
        -Title "Users who viewed reports ($($viewReportUsers.Count))" -TitleBold -TitleSize 14 `
        -TableName 'tblViewReport' -TableStyle Medium2 -AutoSize -PassThru | Out-Null
    Set-TitleBar $pkg.Workbook.Worksheets['ViewReport Users'] 1

    # --- 5) PBIUsage (raw detail) -----------------------------------------
    $flat | Export-Excel -ExcelPackage $pkg -WorksheetName 'PBIUsage' `
        -TableName 'tblPBIUsage' -TableStyle Medium2 -AutoSize -BoldTopRow -PassThru | Out-Null
    $pkg.Workbook.Worksheets['PBIUsage'].View.FreezePanes(2, 1)

    Close-ExcelPackage $pkg
    Write-Host "Exported workbook: $OutputPath" -ForegroundColor Green
    Write-Host ("  Sheets: Overview | Activity Breakdown | CreateReport Users ({0}) | ViewReport Users ({1}) | PBIUsage ({2} rows)" -f `
        $createReportUsers.Count, $viewReportUsers.Count, $flat.Count) -ForegroundColor Green
} else {
    $csv = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
    $flat | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Warning "ImportExcel unavailable. Wrote CSV instead: $csv"
    Write-Host "(Open the CSV in Excel and 'Save As' .xlsx if you need the workbook format.)"
}
