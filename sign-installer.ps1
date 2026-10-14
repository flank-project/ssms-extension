$ErrorActionPreference = "Stop"

$repo = $PSScriptRoot

$signtool = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe"
$dlib = "C:\artifact-signing\Microsoft.ArtifactSigning.Client\bin\x64\Azure.CodeSigning.Dlib.dll"

$metadata = "$repo\installer\artifact-signing.json"
# The installer build-installer.ps1 made for the version in VERSION.
$version = (Get-Content "$repo\VERSION" -Raw).Trim()
$installer = "$repo\installer\output\Flank-SSMS-Setup-$version.exe"
if (-not (Test-Path $installer)) {
    throw "$installer doesn't exist. Run build-installer.ps1 first."
}

Write-Host "Signing Flank installer..."

& $signtool sign `
    /v /debug `
    /fd SHA256 `
    /tr "http://timestamp.acs.microsoft.com" `
    /td SHA256 `
    /dlib $dlib `
    /dmdf $metadata `
    $installer

if ($LASTEXITCODE -ne 0) {
    throw "Signing failed with exit code $LASTEXITCODE"
}

Write-Host "Successfully signed:"
Write-Host $installer