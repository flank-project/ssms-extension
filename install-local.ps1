# Builds Flank, closes SSMS, removes any previous install, installs the new
# build and reopens SSMS. Run from an elevated shell.
#
#   -Unattended   Never prompt. If SSMS doesn't close within -CloseTimeout
#                 seconds (e.g. it's showing a save prompt), force it closed.
#                 Waits for the new SSMS window before exiting. For scripted
#                 and Claude-driven runs.
#   -NoLaunch     Don't reopen SSMS after installing.
param(
    [switch]$Unattended,
    [switch]$NoLaunch,
    [int]$CloseTimeout = 15
)

$ErrorActionPreference = "Stop"

$repo = $PSScriptRoot
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

# Build before touching SSMS, so a failed build leaves the current session alone.
Write-Host "Building Flank..."
& $msbuild "$repo\Flank.SsmsExtension\Flank.SsmsExtension.csproj" `
    /t:Rebuild `
    /p:Configuration=Debug `
    /v:minimal
if ($LASTEXITCODE -ne 0) {
    throw "Build failed."
}

# Ask SSMS to close like clicking X does, so it can offer to save open files.
$running = @(Get-Process SSMS -ErrorAction SilentlyContinue)
if ($running) {
    Write-Host "Closing SSMS..."

    # A modal dialog (e.g. Connect, shown at startup) disables the main window,
    # which then ignores the close request. Close those dialogs first, as if
    # their X were clicked. Done once, before asking SSMS to close, so it never
    # cancels the save prompt that closing SSMS can raise.
    Add-Type -Namespace Flank -Name Win32 -MemberDefinition @'
public delegate bool EnumProc(System.IntPtr hwnd, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern System.IntPtr GetWindow(System.IntPtr hwnd, uint cmd);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool IsWindowVisible(System.IntPtr hwnd);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool PostMessage(System.IntPtr hwnd, uint msg, System.IntPtr wParam, System.IntPtr lParam);
'@
    $GW_OWNER = 4
    $WM_CLOSE = 0x0010
    foreach ($p in $running) {
        $main = $p.MainWindowHandle
        if ($main -eq [IntPtr]::Zero) { continue }
        [void][Flank.Win32]::EnumWindows({
            param($hwnd, $lParam)
            if ([Flank.Win32]::IsWindowVisible($hwnd) -and [Flank.Win32]::GetWindow($hwnd, $GW_OWNER) -eq $main) {
                [void][Flank.Win32]::PostMessage($hwnd, $WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
            }
            $true
        }, [IntPtr]::Zero)
    }
    Start-Sleep -Seconds 1

    foreach ($p in $running) { [void]$p.CloseMainWindow() }

    $deadline = (Get-Date).AddSeconds($CloseTimeout)
    $warned = $false
    while (Get-Process SSMS -ErrorAction SilentlyContinue) {
        if ((Get-Date) -ge $deadline) {
            if ($Unattended) {
                Write-Host "SSMS didn't close within $CloseTimeout seconds; forcing it closed."
                break
            }
            if (-not $warned) {
                Write-Host "SSMS is still open. If it's asking to save changes, answer it there."
                $warned = $true
            }
            $answer = Read-Host "Press Enter to keep waiting, or type F to force SSMS closed (unsaved work is lost)"
            if ($answer -match '^[Ff]') { break }
            $deadline = (Get-Date).AddSeconds($CloseTimeout)
        }
        Start-Sleep -Milliseconds 500
    }

    Stop-Process -Name SSMS -Force -ErrorAction SilentlyContinue
    Wait-Process -Name SSMS -Timeout 30 -ErrorAction SilentlyContinue
}

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

# SSMS is a GUI app, so wait explicitly or /setup can still be running at launch.
Write-Host "Registering extension..."
$setup = Start-Process $ssms -ArgumentList "/setup" -Wait -PassThru
if ($setup.ExitCode -ne 0) {
    throw "SSMS /setup failed with exit code $($setup.ExitCode)."
}

if ($NoLaunch) {
    Write-Host "Done (SSMS not launched)."
    return
}

Write-Host "Launching SSMS..."
$proc = Start-Process $ssms -PassThru

if ($Unattended) {
    Write-Host "Waiting for the SSMS window..."
    $deadline = (Get-Date).AddSeconds(120)
    do {
        Start-Sleep -Seconds 1
        $proc.Refresh()
        if ($proc.HasExited) { throw "SSMS exited during startup (exit code $($proc.ExitCode))." }
    } while ($proc.MainWindowHandle -eq 0 -and (Get-Date) -lt $deadline)

    if ($proc.MainWindowHandle -eq 0) {
        throw "SSMS window didn't appear within 120 seconds."
    }
    Write-Host "SSMS is up (pid $($proc.Id))."
}

Write-Host "Done."
