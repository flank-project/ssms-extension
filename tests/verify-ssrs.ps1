# Renders an SSRS report as CSV and checks it returns data. Prints the
# columns and row count; exits non-zero if rendering fails or returns too
# few rows.
#
#   .\tests\verify-ssrs.ps1 -Report "/FlankTests/My Report"
#   .\tests\verify-ssrs.ps1 -Report "http://localhost/Reports/report/FlankTests/My%20Report"
#   .\tests\verify-ssrs.ps1 -Report "/FlankTests/Sproc" -Parameters "League_Id=1;Name=x"
#
# -Report takes a catalog path or the portal link Flank shows after creating
# a report. Uses the current Windows credentials.
param(
    [Parameter(Mandatory)][string]$Report,
    [string]$Parameters = "",
    [string]$ReportServer = "http://localhost/ReportServer",
    [int]$MinRows = 1
)

$ErrorActionPreference = "Stop"

# Portal links look like http://host/Reports/report/Folder/Name.
if ($Report -match '^https?://[^/]+/[^/]+/report(/.*)$') {
    $Report = [Uri]::UnescapeDataString($Matches[1])
}

$query = "?" + [Uri]::EscapeDataString($Report).Replace("%2F", "/") + "&rs:Format=CSV"
foreach ($pair in $Parameters -split ';' | Where-Object { $_ }) {
    $name, $value = $pair -split '=', 2
    $query += "&" + [Uri]::EscapeDataString($name.Trim()) + "=" + [Uri]::EscapeDataString($value)
}

try {
    $response = Invoke-WebRequest -Uri ($ReportServer.TrimEnd('/') + $query) -UseDefaultCredentials -UseBasicParsing
}
catch {
    $detail = $_.ErrorDetails.Message
    if ($detail -match '<li>([^<]+)</li>') { $detail = $Matches[1] }
    throw "Rendering '$Report' failed: $($_.Exception.Message) $detail"
}

# CSV render is UTF-8 with a BOM.
$csv = [Text.Encoding]::UTF8.GetString($response.RawContentStream.ToArray()).TrimStart([char]0xFEFF)
$rows = @($csv | ConvertFrom-Csv)
$columns = if ($rows.Count -gt 0) { $rows[0].PSObject.Properties.Name } else { @(($csv -split "`r?`n")[0] -split ',') }

Write-Host "${Report}: $($rows.Count) rows; columns: $($columns -join ', ')"

if ($rows.Count -lt $MinRows) {
    throw "Expected at least $MinRows row(s), got $($rows.Count)."
}

Write-Host "OK"
