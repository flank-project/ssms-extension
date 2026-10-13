# Tests

Three layers, cheapest first.

| Layer | What | Run | Time |
|---|---|---|---|
| Unit | Workbook M code, report definitions (RDL), type mapping. No network. | `dotnet test --project Flank.Tests --filter-not-trait "Category=Integration"` | ~15s |
| Integration | Everything after SSMS, for real: generate a workbook and have Excel refresh it; discover columns/parameters, deploy a report to SSRS and render it. | `dotnet test --project Flank.Tests --filter-trait "Category=Integration"` | ~3–5 min |
| SSMS smoke | The SSMS side: menus, reading the query/connection/procedure, dialogs. Driven by Claude with computer use. | `/smoke-test <scenario>` (see `.claude/skills/smoke-test`) | minutes per scenario |

`dotnet test --project Flank.Tests` with no filter runs both test layers.

## Integration test setup

- **Excel** with saved credentials for the test database: open any Flank
  workbook once, Refresh All, and sign in. Excel runs hidden in the tests and
  can't show a prompt.
- **SSRS** at `http://localhost/ReportServer` with a shared data source
  (`/HardcodedTest`) that points at the test database. Tests deploy into
  `/FlankTests`, which they empty at the start of each run.
- **`FLANK_TEST_SQL`**: a connection string for the test database, used to
  discover columns and parameters. Tests that need it are skipped without it.

Optional overrides: `FLANK_TEST_REPORT_SERVER`, `FLANK_TEST_REPORT_PORTAL`,
`FLANK_TEST_REPORT_DATA_SOURCE`.

## Checker scripts

Used by the integration tests and the SSMS smoke tests, so both check the
same thing:

- `verify-excel.ps1 -Path <xlsx>`: opens the workbook in hidden Excel, runs
  Refresh All, and fails unless the query loads rows.
- `verify-ssrs.ps1 -Report <path or portal link> [-Parameters "Name=value;..."]`:
  renders the report as CSV and fails unless it returns rows.

The console apps (`Flank.TestExcelApp`, `Flank.TestSsrsApp`) are still there
for manual poking.
