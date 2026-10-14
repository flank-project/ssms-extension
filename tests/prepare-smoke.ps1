# Resets the places SSMS smoke tests write to:
#   - C:\temp\flank-smoke (workbooks), emptied
#   - SSRS folder /FlankSmoke, emptied (separate from /FlankTests, which the
#     integration tests own)
param(
    [string]$OutputDir = "C:\temp\flank-smoke",
    [string]$ReportServer = "http://localhost/ReportServer",
    [string]$ReportFolder = "/FlankSmoke"
)

$ErrorActionPreference = "Stop"

Remove-Item $OutputDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item $OutputDir -ItemType Directory -Force | Out-Null
Write-Host "Output folder: $OutputDir"

$rs = New-WebServiceProxy -Uri "$($ReportServer.TrimEnd('/'))/ReportService2010.asmx?wsdl" -UseDefaultCredential
if ($rs.ListChildren("/", $false) | Where-Object Path -eq $ReportFolder) {
    $rs.DeleteItem($ReportFolder)
}
[void]$rs.CreateFolder($ReportFolder.TrimStart('/'), "/", $null)
Write-Host "SSRS folder: $ReportFolder"
