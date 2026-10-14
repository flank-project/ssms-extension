---
name: smoke-test
description: Build, install and smoke-test the Flank SSMS extension end to end, driving SSMS with computer use. Scenarios - menus, query-excel, query-ssrs, sproc-excel, sproc-ssrs, or all. Use when asked to smoke-test, run, or check that the extension works in SSMS.
---

# Smoke test

Tests what only SSMS can: the menus, reading the query and connection from
SSMS, and the dialogs. What the workflows produce is checked with the same
scripts the integration tests use (`tests/verify-*.ps1`), not by looking at
Excel or SSRS. See `tests/README.md` for the other test layers.

## Scenarios

The argument picks the scenario; default `menus`.

| Scenario | Checks |
|---|---|
| `menus` | Query editor and stored procedure menus have Make Self-Serve |
| `query-excel` | Query window > Excel: SSMS saves a workbook, and it refreshes with rows |
| `query-ssrs` | Query window > SSRS: SSMS creates a report, and it renders with rows |
| `sproc-excel` | Stored procedure > Excel: workbook for `EXEC` refreshes with rows; with parameters, the user types values above the table and refreshes |
| `sproc-ssrs` | Stored procedure > SSRS: SSMS creates a report, and it renders with rows |
| `all` | All of the above, in that order (before a release). One build and install. |

## 1. Set up (once per run)

From an elevated PowerShell at the repo root (Claude's shell already is):

```powershell
.\tests\prepare-smoke.ps1        # empties C:\temp\flank-smoke and SSRS /FlankSmoke
.\install-local.ps1 -Unattended  # build, reinstall, relaunch SSMS
```

`-Unattended` never prompts: if SSMS doesn't close gracefully within 15s
(usually a save prompt), it is force-closed. It returns once the new SSMS
window exists and exits non-zero if build, setup or launch fails.

Request computer-use access to "SQL Server Management Studio 22" (and
"Excel" for the Excel scenarios), then screenshot.

## 2. How to drive SSMS reliably

- **Keyboard over mouse.** SSMS sometimes ignores mouse clicks (seen on
  Connect, New Query and menu items). Click a dialog's title bar, then
  Return for its default button. Open menus with shift+F10, move with arrow
  keys, and zoom to confirm the highlighted item before pressing Return:
  highlight movement isn't consistent (opening a submenu with Right
  sometimes highlights its first item and sometimes doesn't).
- **Text goes through the clipboard.** Never use the computer-use `type`
  action: without the clipboardWrite grant it pastes whatever the user last
  copied. Run `Set-Clipboard -Value "..."` in PowerShell, then ctrl+a,
  ctrl+v, and zoom to check before continuing.
- **Drop-down lists** (e.g. the SSRS Folder) ignore paste: press one key per
  action to type-ahead (`slash`, `shift+f`, `l`, ...).
- **Bringing SSMS to the front** (e.g. after Excel opened): don't call
  `open_application`, it starts a second SSMS. Run
  `(New-Object -ComObject WScript.Shell).AppActivate((Get-Process SSMS).Id)`.
- **Don't click an already-selected tree node**: a second click starts a
  rename. If an edit box appears, press Escape at once.
- **Recovered Files dialog**: appears if SSMS was force-closed with a query
  open. Click **Do Not Recover** only if the files listed are your own test
  queries; otherwise stop and ask the user.
- Only run read-only SQL.

## 3. Common start

1. Connect dialog: the saved sandbox connection is pre-filled. Click its
   title bar, Return.

For the query scenarios:

2. ctrl+n for a new query window.
3. Put the query on the clipboard and paste it into the editor:
   ```powershell
   Set-Clipboard -Value "-- Flank smoke test`r`nSELECT id,`r`n       name`r`nFROM leagues"
   ```
4. Open the editor's menu: shift+F10, Home (Make Self-Serve), Right to open
   the submenu.

For the stored procedure scenarios:

2. In Object Explorer, expand Databases > fomf-ss-sandbox-db >
   Programmability > Stored Procedures. Click a node's expander, or
   double-click its label; the tree can scroll as it expands, so screenshot
   before each click.
3. Click the procedure to select it, then shift+F10. Make Self-Serve is
   the first item: Down once, then Right opens the submenu.

Procedures used (all read-only):

- `dbo.get_leagues`: no parameters; returns `id, name` (5 rows).
- `dbo.test_multiset_n`: takes `@N` and two OUTPUT parameters; returns three
  result sets (TOP @N, empty, TOP 2×@N). Only the first is used; the Run
  dialog says so and shows all three as tabs. Excel workbooks for it get an
  extra one-time "Native Database Query" approval (click **Run**). SSRS:
  `verify-ssrs.ps1` gives 2 rows with the default N=2, 3 with
  `-Parameters "N=3"`.
- `dbo.sgm_ListTeamsByLeague`: takes `@League_Id`; returns `team_name,
  team_id, league_id`. League 1 and 2 have 2 teams each, league 3 has 1.
  It's further down the list: scroll Object Explorer down to reach it.

## 4. Scenarios

### menus

The submenu shows **Excel (Power Query)** and **SSRS Report**. Escape twice.

Then in Object Explorer, expand Databases > (database) > Programmability >
Stored Procedures (double-click a folder's label to expand it), click a
procedure, and shift+F10: **Make Self-Serve** is the first item, with a
separator under it. Open it and zoom: both items have icons (Excel's and
the SSRS chart, as in the query editor). The Stored Procedures folder's own
menu should not have it.

### query-excel

1. `Set-Clipboard -Value "C:\temp\flank-smoke\query-excel.xlsx"`
2. In the submenu, highlight **Excel (Power Query)** and press Return.
3. Save dialog: ctrl+a, ctrl+v, zoom to check the path, Return.
4. Flank generates the workbook and opens it in Excel (with an "External
   Data Connections have been disabled" bar: expected). Close it with
   ctrl+w; there's nothing to save.
5. Check it:
   ```powershell
   .\tests\verify-excel.ps1 -Path C:\temp\flank-smoke\query-excel.xlsx
   ```
   Pass: `OK`, with 5 rows and columns `id, name`, and a `Last refreshed:` time.

### query-ssrs

1. Bring SSMS to the front if needed; click in the editor and reopen the
   submenu (section 3, step 4).
2. Highlight **SSRS Report**, Return.
3. **Connect to SSRS**: defaults are `http://localhost/ReportServer` and
   `http://localhost/Reports`. Return (Continue).
4. **Create SSRS Report**: focus starts on Folder. It's pre-selected with the
   last folder used (Flank remembers it in HKCU\Software\Flank\Ssrs);
   otherwise type-ahead `/FlankS` one key per action to select `/FlankSmoke`. Tab, Tab to Report Name;
   `Set-Clipboard -Value "Query smoke"`, ctrl+a, ctrl+v. Zoom to check
   Folder, Shared Data Source (`/HardcodedTest`) and Report Name, then
   Return (Create).
5. **SSRS Report Created**: the link is selected; ctrl+c copies it. Escape
   closes the dialog.
6. Check it:
   ```powershell
   .\tests\verify-ssrs.ps1 -Report (Get-Clipboard)
   ```
   Pass: `OK`, with 5 rows and columns `Value_id, Value_name`.

### sproc-excel

1. `Set-Clipboard -Value "C:\temp\flank-smoke\sproc-excel.xlsx"`
2. On `dbo.get_leagues`, open the submenu, highlight **Excel (Power
   Query)**, Return.
3. **Run** dialog ("This procedure has no parameters"): Return runs it.
   Pass: 5 rows in the grid. Return again (Create workbook).
4. Save dialog (default name `get_leagues`): ctrl+a, ctrl+v, zoom, Return.
5. Excel opens the workbook, already showing the 5 rows and a "Last
   refreshed" time; close it with ctrl+w.
6. Check it:
   ```powershell
   .\tests\verify-excel.ps1 -Path C:\temp\flank-smoke\sproc-excel.xlsx
   ```
   Pass: `OK`, with 5 rows and columns `id, name`, and a `Last refreshed:` time.
7. Parameters. `Set-Clipboard -Value "C:\temp\flank-smoke\sproc-excel-params.xlsx"`,
   bring SSMS back, and run **Excel (Power Query)** on
   `dbo.sgm_ListTeamsByLeague`. In the Run dialog, click the
   `@League_Id (int)` box, type `2`, Return (Run). Pass: Team 1 and Team 2
   in the grid. Return (Create workbook); save to that path.
8. Excel opens it. Check the layout: note in row 1, "Last refreshed" with
   the run time in row 2, `League_Id` / `2` in A4:B4, then the table from
   row 6 already showing Team 1 and Team 2. Set the value to 1 with the
   keyboard (Excel ignores some clicks too): ctrl+g, `b`, `4`, Return, then
   `1`, Return. Refresh All with ctrl+alt+F5.
   - **Security notice** ("connects to an external data source"): click
     **OK**. Return picks Cancel.
   - **Privacy levels** (first refresh of a parameter workbook, once per
     file per user): Tab until the "Ignore Privacy Levels checks for this
     file" checkbox has focus (zoom to check; focus order varies), Space,
     Return to Save.
   Pass: Team 3 and Team 5, and `League_Id` / `1` are still in A4:B4 (not
   shifted right).
9. Set B4 to `2` the same way and refresh. Pass: Team 1 and Team 2.
10. ctrl+s (keeps the privacy choice), ctrl+w. Then check it unattended:
    ```powershell
    .\tests\verify-excel.ps1 -Path C:\temp\flank-smoke\sproc-excel-params.xlsx -Parameters "League_Id=3"
    ```
    Pass: `OK`, 1 row. (A parameter workbook that hasn't had the privacy
    prompt answered and saved can't refresh hidden; the checker reports
    that the query didn't load.)

### sproc-ssrs

1. On `dbo.get_leagues`, open the submenu, highlight **SSRS Report**, Return.
2. **Run** dialog: Return runs it (5 rows), Return again (Continue).
3. **Connect to SSRS**: Return (Continue).
4. **Create SSRS Report**: type-ahead `/FlankS` to select `/FlankSmoke`.
   Report Name defaults to `get_leagues`, and a read-only Stored Procedure
   field shows `[dbo].[get_leagues]` (there's no SQL box). Zoom to check,
   then Return (Create).
5. **SSRS Report Created**: ctrl+c copies the link; Escape.
6. Check it:
   ```powershell
   .\tests\verify-ssrs.ps1 -Report (Get-Clipboard)
   ```
   Pass: `OK`, with 5 rows and columns `Value_id, Value_name`.
7. Parameters: on `dbo.sgm_ListTeamsByLeague`, **SSRS Report**; in the Run
   dialog type `2` for `@League_Id`, Return (Run), Return (Continue), then
   the SSRS dialogs as above. The entered value becomes the report
   parameter's default. Check it renders without passing a value, and that
   a value can still be passed:
   ```powershell
   .\tests\verify-ssrs.ps1 -Report (Get-Clipboard)                            # 2 rows (league 2)
   .\tests\verify-ssrs.ps1 -Report (Get-Clipboard) -Parameters "League_Id=3"  # 1 row
   ```

## 5. Report

A table of scenario → pass/fail, with the checker output, and a screenshot
for any failure. Tell the user the clipboard now holds test values. Leave
SSMS open.
