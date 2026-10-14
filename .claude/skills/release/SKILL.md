---
name: release
description: Cut a Flank release - bump VERSION, run the tests, build and sign the installer, smoke-test it in SSMS, and draft the GitHub release. Use when asked to release, ship or cut a version (e.g. "release 0.2.6").
---

# Release

The argument is the new version, `X.Y.Z` (e.g. `0.2.6`). If none is given,
ask; suggest the next patch version after the one in `VERSION`.

The version lives only in `VERSION`. `build-installer.ps1` stamps it on the
installer (name and version) and the DLLs' file version;
`sign-installer.ps1` signs that installer. The README download link is the
one other place it appears.

Publishing the release and merging the PR are the user's call: stop at a
draft and an open PR, and ask.

## 1. Check

- On `main`, clean, and up to date with `origin/main` (`git status`,
  `git fetch`, `git log origin/main..` and `..origin/main` both empty).
- The new version is higher than `VERSION`, and `vX.Y.Z` isn't already a tag
  or release (`gh release view vX.Y.Z` should fail).

Stop and tell the user if any of these fails.

## 2. Test

```powershell
$env:FLANK_TEST_SQL = [Environment]::GetEnvironmentVariable('FLANK_TEST_SQL', 'User')
dotnet test --project Flank.Tests
```

All must pass with none skipped (integration tests need SSRS running locally
and `FLANK_TEST_SQL` set; see `tests/README.md`). Never print
`FLANK_TEST_SQL`: it contains a password.

## 3. Bump

```powershell
git checkout -b release-vX.Y.Z
Set-Content VERSION "X.Y.Z"
```

In `README.md`, update the download line (both the file name and the
`/download/vX.Y.Z/` path) to the new version. Commit both files as
`Release X.Y.Z`.

## 4. Build and sign

```powershell
.\build-installer.ps1
.\sign-installer.ps1
```

`build-installer.ps1` ends with `Done: ...\Flank-SSMS-Setup-X.Y.Z.exe`.
Signing uses Azure Artifact Signing. If it fails with an authentication
error, ask the user to sign in to Azure (e.g. `az login`) and run it again;
don't try to work around it.

Check the signature:

```powershell
Get-AuthenticodeSignature installer\output\Flank-SSMS-Setup-X.Y.Z.exe |
    Select-Object Status, @{n='Signer'; e={ $_.SignerCertificate.Subject }}
```

Pass: `Valid`, signer `CN=Flank Technolgies Inc` (the misspelling is in the
certificate).

## 5. Smoke-test the installer

This replaces whatever Flank build is installed in SSMS with the release.
The installer refuses to run while SSMS is open, so close it first. If the
user has unsaved work in SSMS, ask before closing it.

```powershell
Get-Process SSMS -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow() | Out-Null }
# Wait for it to exit; then:
Start-Process installer\output\Flank-SSMS-Setup-X.Y.Z.exe -ArgumentList '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART' -Wait
```

Check Apps & features / the uninstall key shows the new version:

```powershell
(Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{D5B81B12-1A97-4C4A-9B77-6A996793EBC1}_is1',
                  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{D5B81B12-1A97-4C4A-9B77-6A996793EBC1}_is1' -ErrorAction SilentlyContinue).DisplayVersion
```

Then start SSMS and run the `menus`, `query-excel`, `query-ssrs` and
`sproc-excel` scenarios of the `smoke-test` skill (one of each output, from
both the query editor and Object Explorer), skipping its set-up step (no
`install-local.ps1`: the point is to test the installer). Connect SSMS to
the sandbox by its real name, not `localhost`.

## 6. Draft the release

```powershell
git push -u origin release-vX.Y.Z
gh pr create --base main --title "Release X.Y.Z" --body "<what's in it>"
gh release create vX.Y.Z installer\output\Flank-SSMS-Setup-X.Y.Z.exe `
    --draft --target release-vX.Y.Z --title vX.Y.Z --notes "<notes>"
```

Notes: a short list of user-visible changes since the last release, from
`git log --oneline <previous tag>..HEAD` and the merged PR titles. Plain
language, no commit hashes.

## 7. Hand over

Tell the user: the draft release link, the PR link, test and smoke results,
and the signature check. Ask them to review the draft. Publish
(`gh release edit vX.Y.Z --draft=false`) and merge the PR only when they
say so; publishing creates the tag.
