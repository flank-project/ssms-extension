# Builds Flank (Release) and the installer for the version in VERSION:
# installer\output\Flank-SSMS-Setup-<version>.exe. Sign it with
# sign-installer.ps1.
#
# The version lives only in VERSION. It becomes the installer's version and
# file name, and the DLLs' file version (Properties > Details in Explorer).
$ErrorActionPreference = "Stop"

$repo = $PSScriptRoot
$version = (Get-Content "$repo\VERSION" -Raw).Trim()
if ($version -notmatch '^\d+\.\d+\.\d+$') {
    throw "VERSION must look like 1.2.3, not '$version'."
}

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"

# Find MSBuild: PATH first, then vswhere (same as install-local.ps1).
$msbuild = (Get-Command msbuild -ErrorAction SilentlyContinue).Source
if (-not $msbuild -and (Test-Path $vswhere)) {
    $msbuild = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild `
        -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
}
if (-not $msbuild) {
    throw "MSBuild could not be found."
}

$iscc = @(
    "C:\Program Files\Inno Setup 7\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    throw "Inno Setup (ISCC.exe) could not be found."
}

Write-Host "Building Flank $version (Release)..."
# FileVersion only: the assembly version stays fixed, so nothing that binds
# to these strong-named assemblies changes between releases.
& $msbuild "$repo\Flank.SsmsExtension\Flank.SsmsExtension.csproj" `
    /t:Rebuild `
    /p:Configuration=Release `
    /p:FileVersion=$version.0 `
    /p:InformationalVersion=$version `
    /v:minimal
if ($LASTEXITCODE -ne 0) {
    throw "Build failed."
}

Write-Host "Building installer..."
& $iscc "/DMyAppVersion=$version" "$repo\installer\flank.iss"
if ($LASTEXITCODE -ne 0) {
    throw "Installer build failed."
}

$installer = "$repo\installer\output\Flank-SSMS-Setup-$version.exe"
if (-not (Test-Path $installer)) {
    throw "Expected $installer, but it wasn't created."
}

Write-Host "Done: $installer"
