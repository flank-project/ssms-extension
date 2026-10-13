# Opens a Flank workbook in Excel, runs Refresh All, and checks the Power
# Query loaded data. Prints the column names and row count; exits non-zero
# if the refresh fails or loads no rows.
#
#   .\tests\verify-excel.ps1 -Path C:\temp\report.xlsx [-MinRows 1] [-Timeout 300]
#
# Excel runs hidden. A cold start (open + first refresh) can take over a
# minute. If Excel would need to prompt for database credentials, the refresh
# can't complete unattended: it times out and the script fails.
param(
    [Parameter(Mandatory)][string]$Path,
    [int]$MinRows = 1,
    [int]$Timeout = 300
)

$ErrorActionPreference = "Stop"

$Path = (Resolve-Path $Path).Path

# Run Excel in a job so a hung refresh (e.g. a hidden credentials prompt)
# can be killed by the timeout instead of hanging the caller.
$job = Start-Job -ArgumentList $Path -ScriptBlock {
    param($path)
    $ErrorActionPreference = "Stop"
    $excel = New-Object -ComObject Excel.Application
    try {
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false

        $workbook = $excel.Workbooks.Open($path, 0, $true)
        try {
            # Power Query connections refresh in the background by default.
            # Turn that off so RefreshAll blocks until the data has loaded
            # (CalculateUntilAsyncQueriesDone can hang in a hidden instance).
            foreach ($connection in $workbook.Connections) {
                try { $connection.OLEDBConnection.BackgroundQuery = $false } catch { }
            }
            $workbook.RefreshAll()

            $tables = @()
            foreach ($sheet in $workbook.Worksheets) {
                foreach ($table in $sheet.ListObjects) {
                    $columns = @()
                    foreach ($c in $table.ListColumns) { $columns += $c.Name }
                    $rows = if ($table.DataBodyRange) { $table.DataBodyRange.Rows.Count } else { 0 }
                    $tables += [pscustomobject]@{ Sheet = $sheet.Name; Table = $table.Name; Columns = $columns; Rows = $rows }
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

# Excel started for this job; anything else running is the user's and is left alone.
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)

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

if ($tables.Count -eq 0) {
    throw "The workbook has no tables after refresh."
}

foreach ($t in $tables) {
    Write-Host "$($t.Sheet)/$($t.Table): $($t.Rows) rows; columns: $($t.Columns -join ', ')"
}

$loaded = ($tables | Measure-Object -Property Rows -Sum).Sum
if ($loaded -lt $MinRows) {
    throw "Expected at least $MinRows row(s) after refresh, got $loaded."
}

Write-Host "OK"
