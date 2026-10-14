---
name: smoke-test
description: Build, install and smoke-test the Flank SSMS extension end to end, driving SSMS with computer use. Scenarios - menus, query-excel, query-ssrs, sproc-excel, sproc-ssrs, access, large, or all. Use when asked to smoke-test, run, or check that the extension works in SSMS.
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
| `query-excel` | Query window > Excel, skipping access setup: SSMS saves a workbook, and it refreshes with rows |
| `query-ssrs` | Query window > SSRS: SSMS creates a report, and it renders with rows |
| `sproc-excel` | Stored procedure > Excel: workbook for `EXEC` refreshes with rows; with parameters, the user types values above the table and refreshes |
| `sproc-ssrs` | Stored procedure > SSRS: SSMS creates a report with blank parameters, and it renders with rows |
| `access` | The Excel wizard's access steps: a new SQL user gets SELECT on a query's tables, an existing one EXECUTE on a query turned into a procedure, and each runs exactly what it was granted |
| `large` | Over 50,000 rows: the preview is capped, the workbook starts with headers only and the first refresh loads everything |
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
- **Text boxes in Flank's dialogs**: a click often doesn't move focus, so
  keys and pastes go nowhere. Reach the box with Tab / arrow keys (zoom for
  the caret), then ctrl+v or single `key` presses. If that fails, set it
  directly: find the box's handle with UI Automation (WinForms controls show
  as panes whose AutomationId is the window handle) and send `WM_SETTEXT`;
  `BM_CLICK` does the same for a button. Both raise the usual events.
- **New Query fails with STG_E_FILEALREADYEXISTS** when an empty
  `%TEMP%\SQLQueryN.sql` from an earlier run is left over. Instead, write the
  query to a file under `C:\temp\flank-smoke` and open it (ctrl+o, alt+n,
  paste the path, Return); it connects with the last connection. Never touch
  the user's own query tabs; "changed externally, reload?" prompts about
  them: answer **No**.
- **Excel's Document Recovery** prompt on close: choose "Yes, I want to view
  these files later" (alt+y, Return). The files may be the user's.
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
  dialog shows all three as tabs, and Create workbook / Continue first shows
  a warning ("returns 3 result sets ... only Result 1 will be included",
  Cancel by default; Left then Return picks OK). Excel workbooks for it get an
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
3. **Create Excel workbook** wizard, "Who will refresh this workbook?":
   **Skip this if end user already has access** is selected. Return (Next)
   goes straight to "Run the query" with focus on Run: Return runs it
   (5 rows), Return again (Create workbook).
4. Save dialog: ctrl+a, ctrl+v, zoom to check the path, Return.
5. Flank generates the workbook and opens it in Excel (with an "External
   Data Connections have been disabled" bar: expected), already showing the
   5 rows. Close it with ctrl+w; there's nothing to save.
6. Check it (the checker clears the preview rows first, so rows only count
   if the refresh loaded them):
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
3. Wizard: Return (Next, with Skip selected). "Run [dbo].[get_leagues]"
   ("No parameters."): Return runs it. Pass: 5 rows in the grid. Return
   again (Create workbook).
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
   `dbo.sgm_ListTeamsByLeague`. Wizard: Return (Next, Skip). Focus starts
   in the `@League_Id (int)` box: press `2`, Return (Run). Pass: Team 1 and
   Team 2 in the grid. Return (Create workbook); save to that path.
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
    0 rows and says so.)

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
   dialog press `2` for `@League_Id` (focus starts there), Return (Run),
   Return (Continue), then the SSRS dialogs as above. The report parameter
   has no default (the run is only a preview): without a value it renders
   no rows, with one it filters:
   ```powershell
   .\tests\verify-ssrs.ps1 -Report (Get-Clipboard)                            # fails: 0 rows
   .\tests\verify-ssrs.ps1 -Report (Get-Clipboard) -Parameters "League_Id=3"  # 1 row
   ```

### access

Creates a throwaway SQL user and procedure on the sandbox (allowed: it's a
test database) and drops them at the end. Uses the Excel wizard from a query
window with `SELECT id, name FROM leagues`.

1. `Set-Clipboard -Value "C:\temp\flank-smoke\access-tables.xlsx"`. Open
   **Excel (Power Query)**. Wizard page 1: Down selects "Someone with a SQL
   Server username and password"; existing users load into the list. Tab,
   Down picks "A new user named:", Tab reaches the name box (zoom for the
   caret); `Set-Clipboard -Value "flank_smoke_user"` first, then ctrl+v (or
   WM_SETTEXT, see section 2). The Password row is shown only for a new
   user. Write down the generated password from a zoom. Return (Next).
2. Page 2, "What can they do?": "Turn the query into a stored procedure"
   (name `[dbo].[flank_report]`) and "Let them read the tables the query
   uses: [dbo].[leagues]". Choose the second (Down). Zoom: the script is
   `CREATE USER [flank_smoke_user] WITH PASSWORD = N'...'` (Azure SQL: no
   LOGIN) and `GRANT SELECT ON OBJECT::[dbo].[leagues]`, each followed by
   `GO`. **Apply** (BM_CLICK if the click doesn't take). Pass: "Done. Next,
   run it as flank_smoke_user to check."; Next is enabled only now.
3. Return (Next): "Run the query as flank_smoke_user". Return runs it.
   Pass: 5 rows. Return (Create workbook), save, close Excel.
4. Check it:
   ```powershell
   .\tests\verify-excel.ps1 -Path C:\temp\flank-smoke\access-tables.xlsx
   .\tests\verify-access.ps1 -User flank_smoke_user -Script "WITH PASSWORD = N'<password>'" `
       -Allowed "SELECT * FROM dbo.leagues" `
       -Denied "SELECT * FROM dbo.sgm_teams", "EXEC dbo.sgm_ListTeamsByLeague @League_Id = 1"
   ```
5. Query → procedure, existing user: `Set-Clipboard -Value
   "C:\temp\flank-smoke\access-proc.xlsx"`, open the wizard again. Page 1:
   Down (SQL), Tab, Tab to the user list, press `f` to type-ahead to
   `flank_smoke_user`. Return. Page 2: keep "Turn the query into a stored
   procedure" and set the name to `[dbo].[flank_smoke_report]`. Pass: no
   CREATE USER, just `CREATE PROCEDURE [dbo].[flank_smoke_report]` (the
   query inside `BEGIN SET NOCOUNT ON; ... END`) and `GRANT EXECUTE` on it.
   Apply, Return, Return (Run as flank_smoke_user, 5 rows), create, save.
6. Check it: `verify-excel.ps1` as above, plus `verify-access.ps1 ...
   -Allowed "EXEC dbo.flank_smoke_report"`. The workbook's query is
   `EXEC [dbo].[flank_smoke_report]`, not the original SQL.
7. A procedure from Object Explorer (`dbo.sgm_ListTeamsByLeague`): page 2
   shows only `GRANT EXECUTE ON OBJECT::[dbo].[sgm_ListTeamsByLeague]`.
8. Clean up in your own query tab:
   ```sql
   DROP PROCEDURE IF EXISTS dbo.flank_smoke_report;
   DROP USER IF EXISTS [flank_smoke_user];
   ```

### large

1. Write `select top 60000 a.object_id, b.name from sys.all_objects a cross
   join sys.all_objects b;` to `C:\temp\flank-smoke\large.sql`, open it,
   and run **Excel (Power Query)** with Skip.
2. Run: pass, "over 50,000 rows. Showing the first 50,000." Create workbook
   shows "The result has more than 50,000 rows, too many to include..."
   (Return). Save as `C:\temp\flank-smoke\large.xlsx`.
3. Excel opens it with headings only and "Last refreshed: never". Close it.
4. `.\tests\verify-excel.ps1 -Path C:\temp\flank-smoke\large.xlsx -MinRows 60000`.
   Pass: 60000 rows.

## 5. Report

A table of scenario → pass/fail, with the checker output, and a screenshot
for any failure. Tell the user the clipboard now holds test values. Leave
SSMS open.
