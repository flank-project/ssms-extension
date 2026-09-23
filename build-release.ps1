$ErrorActionPreference = "Stop"

$repo = "C:\Users\angus\source\repos\flank-project\ssms-extension"
$project = "$repo\Flank.SsmsExtension\Flank.SsmsExtension.csproj"
$iss = "$repo\installer\Flank.iss"

Write-Host "Building Flank (Release)..."
msbuild $project `
    /t:Rebuild `
    /p:Configuration=Release

Write-Host "Building installer..."
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" $iss

Write-Host "Done."