# Signs in as a user created by the Excel wizard's access step and
# checks what it can and can't do. The password is read from -Script (the
# wizard's script, or just "WITH PASSWORD = N'...'"), or the clipboard, and is never
# printed.
#
#   .\tests\verify-access.ps1 -User flank_smoke_user `
#       -Allowed "EXEC dbo.sgm_ListTeamsByLeague @League_Id = 1" `
#       -Denied "SELECT * FROM dbo.sgm_teams"
#
# Exits non-zero if any -Allowed statement fails or any -Denied one succeeds.
param(
    [Parameter(Mandatory)][string]$User,
    [string[]]$Allowed = @(),
    [string[]]$Denied = @(),
    [string]$Script = "",
    [string]$Server = "fomf-sandbox-ssdb.database.windows.net",
    [string]$Database = "fomf-ss-sandbox-db"
)

$ErrorActionPreference = "Stop"

if (-not $Script) { $Script = Get-Clipboard -Raw }
if ($Script -notmatch "WITH PASSWORD = N'((?:[^']|'')*)'") {
    throw "No SQL login password found in the script (pass -Script with the wizard's script, or just WITH PASSWORD = N'...')."
}
$password = $Matches[1].Replace("''", "'")

$builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
$builder["Data Source"] = "tcp:$Server,1433"
$builder["Initial Catalog"] = $Database
$builder["User ID"] = $User
$builder["Password"] = $password
$builder["Encrypt"] = $true
$builder["Connect Timeout"] = 60

$connection = New-Object System.Data.SqlClient.SqlConnection $builder.ConnectionString
$connection.Open()
Write-Host "Signed in as $User"

function Try-Run([string]$sql) {
    try {
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
        $reader = $command.ExecuteReader()
        while ($reader.Read()) { }
        $reader.Close()
        return $null
    }
    catch {
        $inner = $_.Exception.InnerException
        return $(if ($inner) { $inner.Message } else { $_.Exception.Message })
    }
}

$failures = 0
foreach ($sql in $Allowed) {
    $err = Try-Run $sql
    if ($err) { Write-Host "FAIL (should be allowed): $sql -> $err"; $failures++ }
    else { Write-Host "ok, allowed: $sql" }
}
foreach ($sql in $Denied) {
    $err = Try-Run $sql
    if (-not $err) { Write-Host "FAIL (should be denied): $sql"; $failures++ }
    else { Write-Host "ok, denied: $sql" }
}

$connection.Close()

if ($failures -gt 0) { throw "$failures access check(s) failed." }
Write-Host "OK"
