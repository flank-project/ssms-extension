$ErrorActionPreference = "Stop"

$repo = "C:\Users\angus\source\repos\flank-project\ssms-extension"
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"

# Registry key Inno Setup writes for the .exe installer (AppId from installer\Flank.iss).
$installerKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{D5B81B12-1A97-4C4A-9B77-6A996793EBC1}_is1",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{D5B81B12-1A97-4C4A-9B77-6A996793EBC1}_is1"
)

# Find SSMS 22 the same way the installer does: default location first, then vswhere.
$ssmsRoot = "C:\Program Files\Microsoft SQL Server Management Studio 22\Release"
if (-not (Test-Path "$ssmsRoot\Common7\IDE\SSMS.exe") -and (Test-Path $vswhere)) {
    $ssmsRoot = & $vswhere -latest -products Microsoft.VisualStudio.Product.SSMS -property installationPath
}
if (-not $ssmsRoot -or -not (Test-Path "$ssmsRoot\Common7\IDE\SSMS.exe")) {
    throw "SSMS 22 could not be found."
}

$ssms = "$ssmsRoot\Common7\IDE\SSMS.exe"
$dest = "$ssmsRoot\Common7\IDE\Extensions\Flank"
Write-Host "Using SSMS at $ssmsRoot"

# Find MSBuild: PATH first, then vswhere.
$msbuild = (Get-Command msbuild -ErrorAction SilentlyContinue).Source
if (-not $msbuild -and (Test-Path $vswhere)) {
    $msbuild = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild `
        -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
}
if (-not $msbuild) {
    throw "MSBuild could not be found."
}

Write-Host "Building Flank..."
& $msbuild "$repo\Flank.SsmsExtension\Flank.SsmsExtension.csproj" `
    /t:Rebuild `
    /p:Configuration=Debug
if ($LASTEXITCODE -ne 0) {
    throw "Build failed."
}

Write-Host "Closing SSMS..."
Stop-Process -Name SSMS -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# Remove any copy installed by the .exe installer, so Windows doesn't keep
# listing an install whose uninstaller this script is about to delete.
foreach ($key in $installerKeys) {
    if (-not (Test-Path $key)) { continue }

    $uninstaller = (Get-ItemProperty $key).UninstallString -replace '"', ''

    if ($uninstaller -and (Test-Path $uninstaller)) {
        Write-Host "Uninstalling previous .exe install..."
        Start-Process $uninstaller -ArgumentList "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART" -Wait

        # The Inno uninstaller hands off to a temp copy of itself, so wait for
        # the registry key to disappear rather than trusting -Wait alone.
        $deadline = (Get-Date).AddSeconds(60)
        while ((Test-Path $key) -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 500
        }
    }

    if (Test-Path $key) {
        Write-Host "Removing orphaned .exe install entry..."
        Remove-Item $key -Recurse -Force
    }
}

Write-Host "Installing extension..."
Remove-Item $dest -Recurse -Force -ErrorAction SilentlyContinue
New-Item $dest -ItemType Directory -Force | Out-Null

$build = "$repo\Flank.SsmsExtension\bin\Debug\net472"

Copy-Item "$build\Flank.SsmsExtension.dll" $dest
Copy-Item "$build\Flank.SsmsExtension.pkgdef" $dest
Copy-Item "$build\Flank.Excel.dll" $dest
Copy-Item "$build\Flank.Ssrs.dll" $dest

Write-Host "Registering extension..."
& $ssms /setup

Write-Host "Launching SSMS..."
Start-Process $ssms

Write-Host "Done."
