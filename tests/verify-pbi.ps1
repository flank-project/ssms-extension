# Exports a Power BI paginated report as CSV and checks it returns data.
# Prints the columns and row count; exits non-zero if the export fails or
# returns too few rows.
#
#   .\tests\verify-pbi.ps1 -Report "My Report"
#   .\tests\verify-pbi.ps1 -Report "https://app.powerbi.com/groups/<id>/rdlreports/<id>"
#   .\tests\verify-pbi.ps1 -Report "Sproc" -Parameters "League_Id=1;Name=x"
#
# -Report takes a report name in the test workspace or the link Flank shows
# after publishing. Signs in as the test service principal from FLANK_TEST_PBI
# ("TenantId=...;ClientId=...;ClientSecret=...;Workspace=..."). Export needs
# the workspace on Premium or Fabric (trial) capacity.
param(
    [Parameter(Mandatory)][string]$Report,
    [string]$Parameters = "",
    [int]$MinRows = 1
)

$ErrorActionPreference = "Stop"

$settings = [Environment]::GetEnvironmentVariable('FLANK_TEST_PBI')
if (-not $settings) { $settings = [Environment]::GetEnvironmentVariable('FLANK_TEST_PBI', 'User') }
if (-not $settings) { throw "Set FLANK_TEST_PBI (TenantId=...;ClientId=...;ClientSecret=...;Workspace=...)." }

$pbi = @{}
foreach ($pair in $settings -split ';') {
    if ($pair -match '^\s*([^=]+?)\s*=\s*(.*)$') { $pbi[$Matches[1].ToLower()] = $Matches[2].Trim() }
}

$token = (Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($pbi.tenantid)/oauth2/v2.0/token" -Body @{
    grant_type    = 'client_credentials'
    client_id     = $pbi.clientid
    client_secret = $pbi.clientsecret
    scope         = 'https://analysis.windows.net/powerbi/api/.default'
}).access_token
$headers = @{ Authorization = "Bearer $token" }
$api = 'https://api.powerbi.com/v1.0/myorg'

function Invoke-PowerBi([string]$Method, [string]$Path, $Body) {
    $call = @{ Method = $Method; Uri = "$api$Path"; Headers = $headers }
    if ($Body) { $call.ContentType = 'application/json'; $call.Body = ($Body | ConvertTo-Json -Depth 6) }
    try { Invoke-RestMethod @call }
    catch { throw "Power BI $Method $Path failed: $($_.Exception.Message) $($_.ErrorDetails.Message)" }
}

# Links look like https://app.powerbi.com/groups/<workspace>/rdlreports/<report>.
if ($Report -match '/groups/([0-9a-f-]{36})/rdlreports/([0-9a-f-]{36})') {
    $workspaceId, $reportId = $Matches[1], $Matches[2]
    $name = (Invoke-PowerBi GET "/groups/$workspaceId/reports/$reportId").name
}
else {
    $workspaceId = (Invoke-PowerBi GET "/groups?`$filter=name eq '$($pbi.workspace)'").value[0].id
    if (-not $workspaceId) { throw "The service principal can't see workspace '$($pbi.workspace)'." }
    $match = (Invoke-PowerBi GET "/groups/$workspaceId/reports").value |
        Where-Object { $_.reportType -eq 'PaginatedReport' -and $_.name -eq $Report }
    if (-not $match) { throw "No paginated report named '$Report' in '$($pbi.workspace)'." }
    $reportId, $name = $match[0].id, $match[0].name
}

$body = @{ format = 'CSV' }
$values = @($Parameters -split ';' | Where-Object { $_ } | ForEach-Object {
    $n, $v = $_ -split '=', 2
    @{ name = $n.Trim(); value = $v }
})
if ($values.Count -gt 0) { $body.paginatedReportConfiguration = @{ parameterValues = $values } }

$export = Invoke-PowerBi POST "/groups/$workspaceId/reports/$reportId/ExportTo" $body
for ($i = 0; $i -lt 100; $i++) {
    Start-Sleep -Seconds 3
    $status = Invoke-PowerBi GET "/groups/$workspaceId/reports/$reportId/exports/$($export.id)"
    if ($status.status -in 'Succeeded', 'Failed') { break }
}
if ($status.status -ne 'Succeeded') {
    throw "Exporting '$name' failed: $($status.status) $($status.error | ConvertTo-Json -Compress)"
}

$file = Invoke-WebRequest -Uri "$api/groups/$workspaceId/reports/$reportId/exports/$($export.id)/file" -Headers $headers -UseBasicParsing
$csv = [Text.Encoding]::UTF8.GetString($file.RawContentStream.ToArray()).TrimStart([char]0xFEFF)
$rows = @($csv | ConvertFrom-Csv)
$columns = if ($rows.Count -gt 0) { $rows[0].PSObject.Properties.Name } else { @(($csv -split "`r?`n")[0] -split ',') }

Write-Host "${name}: $($rows.Count) rows; columns: $($columns -join ', ')"

if ($rows.Count -lt $MinRows) {
    throw "Expected at least $MinRows row(s), got $($rows.Count)."
}

Write-Host "OK"
