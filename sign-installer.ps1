$ErrorActionPreference = "Stop"

$repo = $PSScriptRoot

$signtool = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe"
$dlib = "C:\artifact-signing\Microsoft.ArtifactSigning.Client\bin\x64\Azure.CodeSigning.Dlib.dll"

$metadata = "$repo\installer\artifact-signing.json"
$installer = "$repo\installer\output\Flank-SSMS-Setup-0.2.1.exe"

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