---
name: smoke-test
description: Build, install and smoke-test the Flank SSMS extension end to end, driving SSMS with computer use. Use when asked to smoke-test, run, or check that the extension works in SSMS.
---

# Smoke test

## 1. Build and install

Run from an elevated PowerShell (Claude's shell already is):

```powershell
.\install-local.ps1 -Unattended
```

`-Unattended` never prompts: if SSMS doesn't close gracefully within 15s
(usually a save prompt), it is force-closed. The script returns once the new
SSMS window exists, and exits non-zero if build, setup or launch fails.

## 2. Drive SSMS

Request computer-use access to "SQL Server Management Studio 22", then
screenshot. The SSMS the script launches can be driven directly on this
machine (the whole desktop session runs elevated). Don't call
`open_application` for SSMS while it's running: it starts a second instance.

Mouse clicks on SSMS are sometimes ignored (seen on the Connect button and
on context-menu items). Prefer the keyboard: click a dialog's title bar,
then Return for its default button; select a tree node with a click, then
shift+F10 for its context menu, arrow keys, and Return.

- **Recovered Files dialog**: appears if SSMS was force-closed with a query
  open. Click **Do Not Recover** only if the files listed are your own test
  queries; otherwise stop and ask the user.
- **Connect dialog**: the saved sandbox connection is pre-filled; click
  **Connect**.
- **New Query**, then enter SQL. Do NOT use the computer-use `type` action:
  it pastes via the clipboard and, without the clipboardWrite grant, pastes
  whatever the user last copied. Instead run
  `Set-Clipboard -Value "SELECT 1 AS smoke_test;"` in PowerShell, then
  click the editor, ctrl+a, ctrl+v, and zoom to confirm the text before F5.
- Only run read-only SQL.

## 3. Check

- The query runs and the status bar says "Query executed successfully".
- Right-click in the query editor: **Make Self-Serve** is present, with
  **Excel (Power Query)** and **SSRS Report** in its submenu.
- Press Escape to close the menus.
- In Object Explorer, expand Databases > (database) > Programmability >
  Stored Procedures and open a procedure's context menu: **Make Self-Serve**
  is present (the folder's own menu should not have it).

Don't click Excel (Power Query) or SSRS Report unless the user asked to test
them: they launch Excel or contact a report server.

Report what was checked, with a screenshot of the submenu if useful. Tell
the user the clipboard now holds the test query.
