# Opens a Flank workbook in Excel, runs Refresh All, and checks the Power
# Query loaded data. Prints the column names and row count; exits non-zero
# if the refresh fails or loads no rows.
#
#   .\tests\verify-excel.ps1 -Path C:\temp\report.xlsx [-MinRows 1] [-Timeout 300]
#   .\tests\verify-excel.ps1 -Path C:\temp\report.xlsx -Parameters "League_Id=1;Name=x"
#
# -Parameters fills in the workbook's parameter cells (labels in column A,
# values in column B, each named Flank_<parameter>) before refreshing. The file is opened read-only and
# never saved.
#
# Excel runs hidden. A cold start (open + first refresh) can take over a
# minute. If Excel would need to prompt for database credentials, the refresh
# can't complete unattended: it times out and the script fails.
param(
    [Parameter(Mandatory)][string]$Path,
    [string]$Parameters = "",
    [int]$MinRows = 1,
    [int]$Timeout = 300
)

$ErrorActionPreference = "Stop"

$Path = (Resolve-Path $Path).Path

$values = [ordered]@{}
foreach ($pair in $Parameters -split ';' | Where-Object { $_ }) {
    $name, $value = $pair -split '=', 2
    $values[$name.Trim()] = $value
}

# Excel processes that exist now are the user's and are left alone.
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)

# Run Excel in a job so a hung refresh (e.g. a hidden credentials prompt)
# can be killed by the timeout instead of hanging the caller.
$job = Start-Job -ArgumentList $Path, $values -ScriptBlock {
    param($path, $values)
    $ErrorActionPreference = "Stop"
    $excel = New-Object -ComObject Excel.Application
    try {
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false

        $workbook = $excel.Workbooks.Open($path, 0, $true)
        try {
            if ($values.Count -gt 0) {
                # Each parameter's value cell is named Flank_<parameter>; its label is beside it in column A.
                $missing = [System.Collections.Generic.List[string]]::new([string[]]$values.Keys)
                foreach ($definedName in $workbook.Names) {
                    if (-not $definedName.Name.StartsWith("Flank_")) { continue }
                    $cell = $definedName.RefersToRange
                    $name = [string]$cell.Offset(0, -1).Value2
                    if ($values.Contains($name)) {
                        $cell.Value2 = $values[$name]
                        [void]$missing.Remove($name)
                    }
                }
                if ($missing.Count -gt 0) { throw "No parameter cell for: $($missing -join ', ')" }
            }

            # Power Query connections refresh in the background by default.
            # Turn that off so RefreshAll blocks until the data has loaded
            # (CalculateUntilAsyncQueriesDone can hang in a hidden instance).
            foreach ($connection in $workbook.Connections) {
                try { $connection.OLEDBConnection.BackgroundQuery = $false } catch { }
            }

            # Workbooks start with the preview's rows. Clear them so rows can
            # only come from this refresh (a failed one leaves the table empty).
            # Only query tables (SourceType 3): a table parameter's rows are input.
            foreach ($sheet in $workbook.Worksheets) {
                foreach ($table in $sheet.ListObjects) {
                    if ($table.Name -ne "LastRefreshed" -and $table.SourceType -eq 3 -and $table.DataBodyRange) {
                        [void]$table.DataBodyRange.ClearContents()
                    }
                }
            }

            $refreshError = $null
            try { $workbook.RefreshAll() } catch { $refreshError = $_.Exception.Message }

            $tables = @()
            foreach ($sheet in $workbook.Worksheets) {
                foreach ($table in $sheet.ListObjects) {
                    $columns = @()
                    foreach ($c in $table.ListColumns) { $columns += $c.Name }
                    # Non-empty rows: cleared rows that didn't reload don't count.
                    $rows = if ($table.DataBodyRange) {
                        [int]$excel.WorksheetFunction.CountA($table.DataBodyRange.Columns.Item(1))
                    } else { 0 }
                    $first = if ($table.DataBodyRange) { $table.DataBodyRange.Cells.Item(1, 1).Text } else { "" }
                    $tables += [pscustomobject]@{ Sheet = $sheet.Name; Table = $table.Name; Columns = $columns; Rows = $rows; First = $first; Error = $refreshError }
                }
            }
            $tables
        }
        finally {
            $workbook.Close($false)
        }
    }
    finally {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}

if (-not (Wait-Job $job -Timeout $Timeout)) {
    Stop-Job $job
    Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $before -notcontains $_.Id } |
        Stop-Process -Force
    Remove-Job $job -Force
    throw "Excel refresh didn't finish within $Timeout seconds. If Excel needs database credentials, open the workbook once, refresh, and save them."
}

$tables = @(Receive-Job $job)
Remove-Job $job

# Quit doesn't always end the hidden instance; stop any new, windowless one.
Get-Process EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $before -notcontains $_.Id -and -not $_.MainWindowTitle } |
    Stop-Process -Force -ErrorAction SilentlyContinue

if ($tables.Count -eq 0) {
    throw "The workbook has no tables after refresh."
}

# The 1-cell "Last refreshed" table isn't data; report it separately.
$stamp = $tables | Where-Object Table -eq "LastRefreshed"
$tables = @($tables | Where-Object Table -ne "LastRefreshed")

# Dropdown choices (the hidden DropdownOptions sheet) aren't the data either.
$tables = @($tables | Where-Object Sheet -ne "DropdownOptions")

foreach ($t in $tables) {
    Write-Host "$($t.Sheet)/$($t.Table): $($t.Rows) rows; columns: $($t.Columns -join ', ')"
}
if ($stamp) {
    Write-Host "Last refreshed: $($stamp.First)"
    if (-not $stamp.First) { throw "The Last refreshed cell is empty after refresh." }
}

if ($tables[0].Error) {
    throw "Refresh failed: $($tables[0].Error)"
}

# A failed refresh leaves the template's placeholder column (and its empty
# row) in place, which would otherwise count as a loaded row.
if ($tables | Where-Object { $_.Columns -match 'Refresh All' }) {
    throw "The query didn't load: the table still has the template's placeholder column."
}

$loaded = ($tables | Measure-Object -Property Rows -Sum).Sum
if ($loaded -lt $MinRows) {
    $hint = if ($values.Count -gt 0) {
        " Workbooks with parameters can't refresh hidden (Power Query asks about privacy levels first); check them in Excel."
    } else { "" }
    throw "Expected at least $MinRows row(s) after refresh, got $loaded.$hint"
}

Write-Host "OK"
