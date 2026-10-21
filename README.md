# Flank

Flank is an SSMS extension that turns SQL queries into self-serve reports.

Right-click a query (or a stored procedure in Object Explorer) and choose:

- **SSRS Report** — deploy the query as a report to an existing SSRS server.
- **Excel (Power Query)** — create an .xlsx workbook that users can refresh themselves.
- **Power BI Paginated Report** — publish the query as a paginated report to a Power BI workspace.

Using Flank? I'd love to hear how — email [anguspmitchell@gmail.com](mailto:anguspmitchell@gmail.com), or use **Make Self-Serve → Feedback...** in SSMS.

<img width="1600" height="682" alt="Flank SSRS SSMS README" src="https://github.com/user-attachments/assets/1f21ef93-ad61-408b-bfa9-9300cef3bd66" />

---

**Jump to:** [Install](#install) · [SSRS Reports](#ssrs-reports) · [Excel (Power Query)](#excel-power-query) · [Power BI Paginated Reports](#power-bi-paginated-reports) · [Reviewing parameters](#reviewing-parameters) · [Uninstall](#uninstall)

---

## Install

Flank currently supports **SSMS 22 on Windows**. For more detailed compatibility information, see [SSRS Reports](#ssrs-reports) and [Excel (Power Query)](#excel-power-query).

Download [Flank-SSMS-Setup-0.2.7.exe](https://github.com/flank-project/ssms-extension/releases/download/v0.2.7/Flank-SSMS-Setup-0.2.7.exe).

Close SSMS, run the installer, then reopen SSMS.

Right-click inside a SQL query, or on a stored procedure in Object Explorer, and you should see **Make Self-Serve**, with options for **SSRS Report** and **Excel (Power Query)**.

## SSRS Reports

Flank can turn a SQL query into a basic SSRS report and deploy it directly to your existing SSRS server.

### Current support

Flank currently supports:

- SSRS with Windows Authentication
- Existing shared data sources
- Selected SQL, or the entire query if nothing is selected
- Stored procedures, from Object Explorer. After you [review the parameters](#reviewing-parameters), Flank asks for values and runs the procedure so you can preview the results; the report's parameters use the names, starting values and blank behaviour you chose
- Basic tabular reports using the columns returned by the query

Not supported yet:

- Multiple result sets: an SSRS dataset only reads a procedure's first result set
- Report formatting and customization
- Creating or configuring SSRS data sources

### Create a report

Right-click a query and choose **Make Self-Serve → SSRS Report**. For a stored procedure, right-click it in Object Explorer (**Programmability → Stored Procedures**) instead.

Flank will ask for two URLs:

- **Report Server URL** — the SSRS web service URL, typically `http://your-server/ReportServer`
- **Report Portal URL** — the SSRS web portal URL, typically `http://your-server/Reports`

If you normally view reports in a browser, the URL you use is the **Report Portal URL**. The **Report Server URL** is usually the same server with `/ReportServer` instead of `/Reports`.

Then choose a folder and shared data source, name the report, and click **Create Report**.

Flank remembers the server URLs, folder and data source from the last report you created (per Windows user) and pre-fills them next time.

## Excel (Power Query)

Flank can turn a SQL query into an .xlsx workbook that end users can refresh themselves using Power Query.

### Current support

Flank currently supports:

- **Desktop Excel** with Power Query
  - End users can "Refresh All" using their own database credentials
- **Native SQL queries**
  - Uses Power Query's "native SQL query" support
  - Includes `SELECT` queries, CTEs, temp tables, variables, and stored procedures that return a result set
- **Stored procedures with parameters**, from Object Explorer
  - The parameters are listed at the top (labels in column A, values in column B, notes in column C). End users change the values and refresh
  - Each value cell only accepts its type: whole numbers in the SQL type's range, numbers, dates, text up to its length, and TRUE/FALSE (a dropdown) for `bit`. Selecting a cell shows its note, and required cells are highlighted while empty
  - Refreshing checks the values again, since pasting gets past Excel's checks, and says which cell is wrong
  - See [Reviewing parameters](#reviewing-parameters)
- A short wizard: who will refresh the workbook and what they can do (see [Giving end users access](#giving-end-users-access)), then a run as that person to preview the results
- The workbook starts with the previewed values and rows, so end users see data straight away
  - Over 50,000 rows, the preview shows the first 50,000 and the workbook starts with just the column headings; the first **Refresh All** loads everything
- Every workbook has one sheet, with a short "how to update" note and a **Last refreshed** time

Not supported yet:

- SQL
  - Parameters for queries from the query editor (stored procedures only)
  - Multiple result sets: Power Query only reads a procedure's first result set, so that's the one the workbook uses. Flank's preview shows them all and warns before creating the workbook; to include the others, split the procedure into one per result set
- Excel
  - Excel for the web
 
### The privacy levels prompt (workbooks with parameters)

The first time someone refreshes a workbook with parameters, Excel asks about **privacy levels**, because the query combines values from the workbook with data from the database. Tick **Ignore Privacy Levels checks for this file**, click **Save**, then save the workbook. Excel remembers the choice for that file, for that Windows user, so each person sees it once.

Flank can't make that choice in advance: Excel stores it encrypted per Windows user.

For a stored procedure with `OUTPUT` parameters, or with a parameter that uses the procedure's default when it's left blank, Excel also asks once to approve the workbook's **native database query** (the procedure call, built so it can leave out blank parameters or declare a variable for each output). Click **Run**.

### Giving end users access

The workbook connects directly to your database through Power Query, so the person refreshing it needs their own database access. Flank never puts your credentials in the workbook. End users don't need Flank installed.

This is part of **Make Self-Serve → Excel (Power Query)**, before anything runs:

1. **Who will refresh this workbook?** Choose **Skip this if end user already has access**, or how they sign in, then an existing database user or a new one:
   - a **SQL Server username and password** (Flank generates the password), or
   - their **Microsoft Entra** account (Azure SQL), or
   - their **Windows / Active Directory** account (SQL Server)
2. **What can they do?** Only what's needed:
   - a stored procedure: `EXECUTE` on it
   - a query: either turn it into a stored procedure and grant `EXECUTE` (the workbook then calls the procedure), or grant `SELECT` on the tables and views it reads. Flank finds those by parsing the query; table names built at run time (dynamic SQL) aren't found

   Flank shows the exact T-SQL and tells you what to send the person and which sign-in option to pick when Excel asks. **Run SQL** runs it; if it fails, Flank says why and you stay on that step.
3. **Run as them.** Flank runs the query or procedure as that user (`EXECUTE AS USER`), so you see what they'll see before sending anything. That needs `IMPERSONATE` permission on them (db_owner has it; granting access alone doesn't give it). If you can't impersonate them, choose **Run as you instead**. Microsoft Entra users can't be impersonated this way, so for them the run is always as you, and the button says so.

Prefer an identity people already have: Entra on Azure SQL, Windows / Active Directory on SQL Server. Use SQL usernames when those aren't available.

One-time setup Flank doesn't do for you:

- **Network access:** the end user's computer must reach the server (for example, on the corporate network or VPN). On Azure SQL, the server's firewall must allow their IP address or network.
- **Server name:** the workbook uses the server name SSMS is connected with. Connect using a name others can reach (e.g. `SQLPROD01`), not `localhost` or `.`; Flank warns if you don't.
- **SQL usernames on SQL Server:** the server must allow SQL Server authentication (Server properties → Security → *SQL Server and Windows Authentication mode*, then restart SQL Server). Azure SQL always allows it; Flank creates the user in the database itself.
- **Microsoft Entra:** the Azure SQL server needs a Microsoft Entra admin (Azure Portal → the SQL server → Microsoft Entra ID), and you must be connected with Entra to create Entra users. The person must already be in your Entra tenant; use their User Principal Name (often, but not always, their email address).

If end users aren't allowed to connect to the database at all, Flank's refreshable Excel output isn't a fit: there's no Flank server or service account between Excel and the database.

## Power BI Paginated Reports

Flank can publish a SQL query or stored procedure as a paginated report in a
Power BI workspace. Paginated reports are SSRS reports hosted in Power BI, so
you get the same basic tabular report as **SSRS Report**, without an SSRS server.

### Current support

Flank currently supports:

- **Azure SQL Database** (servers ending in `.database.windows.net`). Power BI
  reaches other servers only through an on-premises data gateway, which Flank
  doesn't set up yet
- Selected SQL, or the entire query if nothing is selected
- Stored procedures, from Object Explorer, with their parameters as report
  parameters, set up on the [review page](#reviewing-parameters)
- Signing in to Power BI with your work or school account, in the browser or
  with a code you enter in any browser. Flank remembers the sign-in (per
  Windows user), so you sign in once
- How the report reads the database, chosen when you publish:
  - **A SQL login saved on the report.** Flank fills it in from your SSMS
    connection when that uses SQL Server authentication
  - **A shared connection** (Power BI's equivalent of an SSRS shared data
    source): pick an existing one for the database, or create one with a SQL
    login. Flank asks you to sign in once more to allow this

Publishing again with the same name in the same workspace replaces the report.

Not supported yet:

- On-premises SQL Server (needs a data gateway)
- Viewers reading the database as themselves (single sign-on)
- Multiple result sets, report formatting

### Create a report

Right-click a query and choose **Make Self-Serve → Power BI Paginated Report**,
or right-click a stored procedure in Object Explorer. Sign in if Flank asks,
then choose the workspace, name the report, choose how it reads the database,
and click **Publish**.

The workspace needs a Power BI Pro or Premium Per User licence, or Premium/Fabric
capacity, to hold paginated reports.

## Reviewing parameters

Before creating a workbook or report from a stored procedure with parameters,
Flank shows them for review. For each one:

- **Shown as**: the name people see (the Excel label, or the report prompt)
- **Note for the user**: shown next to the Excel cell and when it's selected.
  SSRS and Power BI have nowhere to show it
- **Starts with**: the value it starts with. Flank fills this in from the
  procedure's own default where it declares one (`@n int = 5`)
- **If left blank**:
  - **Required**: a value must be entered
  - **Use procedure default**: the parameter is left out of the call, so the
    procedure's own default applies (only for parameters that have one)
  - **Pass NULL**

## Uninstall

Remove Flank normally from **Windows Settings → Apps → Installed apps**.

Close SSMS before uninstalling.
