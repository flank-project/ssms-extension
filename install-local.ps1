$ErrorActionPreference = "Stop"

$repo = "C:\Users\angus\source\repos\flank-project\ssms-extension"
$dest = "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\Extensions\Flank"
$ssms = "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\SSMS.exe"

Write-Host "Building Flank..."
msbuild "$repo\Flank.SsmsExtension\Flank.SsmsExtension.csproj" `
    /t:Rebuild `
    /p:Configuration=Debug

Write-Host "Closing SSMS..."
Stop-Process -Name SSMS -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

Write-Host "Installing extension..."
Remove-Item $dest -Recurse -Force -ErrorAction SilentlyContinue
New-Item $dest -ItemType Directory -Force | Out-Null

$build = "$repo\Flank.SsmsExtension\bin\Debug\net472"

Copy-Item "$build\Flank.SsmsExtension.dll" $dest
Copy-Item "$build\Flank.SsmsExtension.pkgdef" $dest
Copy-Item "$build\Flank.Excel.dll" $dest

Write-Host "Registering extension..."
& $ssms /setup

Write-Host "Launching SSMS..."
Start-Process $ssms

Write-Host "Done."