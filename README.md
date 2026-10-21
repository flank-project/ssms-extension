# Flank

Flank is an SSMS extension that turns SQL queries into self-serve reports.

Right-click a query (or a stored procedure in Object Explorer) and choose:

- **Excel (Power Query)**
- **SSRS Report**
- **Power BI (Paginated Report) (Beta)**

<img width="1600" height="682" alt="Flank SSRS SSMS README" src="https://github.com/user-attachments/assets/1f21ef93-ad61-408b-bfa9-9300cef3bd66" />

---

**Jump to:** [Install](#install) · [Excel (Power Query)](#excel-power-query) · [SSRS Reports](#ssrs-reports) · [Power BI Paginated Reports](#power-bi-paginated-reports) · [Uninstall](#uninstall)

---

## Install

| Version | Download |
|---|---|
| **0.3.0** | **[⬇ Flank-SSMS-Setup-0.3.0.exe](https://github.com/flank-project/ssms-extension/releases/download/v0.3.0/Flank-SSMS-Setup-0.3.0.exe)** |

Close SSMS, run the installer, then reopen SSMS.

Requires SSMS 22 on Windows. Right-click inside a SQL query, or on a stored procedure in Object Explorer, and you should see **Make Self-Serve**.

## Excel (Power Query)

Flank can turn a SQL query into an .xlsx workbook that end users can refresh themselves using Power Query.

### Current support

- Excel
  - Desktop Excel with Power Query. End users **Refresh All** with their own database credentials; they don't need Flank
  - Parameters at the top of the sheet. Each cell only accepts its type (`bit` is a checkbox), required cells are highlighted, and refreshing checks the values again
  - Dropdowns for parameters, with choices from a stored procedure, a query, or values you type in
  - The workbook opens with the rows from Flank's preview run (up to 50,000; with more, the first **Refresh All** loads them)
- SQL
  - Selected SQL, or the entire query if nothing is selected
  - `SELECT` queries, CTEs, temp tables, variables, and stored procedures that return a result set
  - Parameters in a query (`WHERE league_id = @League_Id`), run as written or turned into a stored procedure
  - Stored procedures, from Object Explorer
- Access
  - Optionally creates a database user for the person refreshing, grants only what they need, and runs the query as them to preview (see [Giving end users access](#giving-end-users-access))

### Not supported yet

- Excel
  - Excel for the web
- SQL
  - Multiple result sets (the workbook uses the first)

### Excel prompts with parameters

The first time someone refreshes a workbook with parameters, Excel asks about **privacy levels**. Tick **Ignore Privacy Levels checks for this file**, click **Save**, then save the workbook. Each person sees it once per file.

Excel may also ask once to approve the workbook's **native database query**. Click **Run**.

### Giving end users access

The workbook connects directly to your database, so the person refreshing it needs their own database access. Flank never puts your credentials in the workbook.

When you create a workbook, Flank asks who will refresh it (a SQL username, a Microsoft Entra account, or a Windows account), shows the T-SQL to give them only what they need (`EXECUTE` on a procedure, or `SELECT` on the tables a query reads), and runs it as them so you see what they'll see. Choose **Skip this if end user already has access** if they already have it.

One-time setup Flank doesn't do for you:

- **Network access:** the end user's computer must reach the server (for example, on the corporate network or VPN). On Azure SQL, the server's firewall must allow their IP address or network.
- **Server name:** the workbook uses the server name SSMS is connected with. Connect using a name others can reach (e.g. `SQLPROD01`), not `localhost` or `.`; Flank warns if you don't.
- **SQL usernames on SQL Server:** the server must allow SQL Server authentication (Server properties → Security → *SQL Server and Windows Authentication mode*, then restart SQL Server).
- **Microsoft Entra:** the Azure SQL server needs a Microsoft Entra admin (Azure Portal → the SQL server → Microsoft Entra ID), and you must be connected with Entra to create Entra users.

## SSRS Reports

Flank can turn a SQL query into a basic SSRS report and deploy it directly to your existing SSRS server.

### Current support

- SSRS
  - Windows Authentication
  - Existing shared data sources
  - Basic tabular reports using the columns returned by the query
  - Report parameters, with dropdowns from a stored procedure, a query, or values you type in
- SQL
  - Selected SQL, or the entire query if nothing is selected
  - Parameters in a query (`WHERE league_id = @League_Id`), run as written or turned into a stored procedure
  - Stored procedures, from Object Explorer

### Not supported yet

- SSRS
  - Creating or configuring SSRS data sources
  - Report formatting and customization
- SQL
  - Multiple result sets

### Create a report

Right-click a query and choose **Make Self-Serve → SSRS Report**. For a stored procedure, right-click it in Object Explorer (**Programmability → Stored Procedures**) instead.

Flank will ask for two URLs:

- **Report Server URL** — the SSRS web service URL, typically `http://your-server/ReportServer`
- **Report Portal URL** — the SSRS web portal URL, typically `http://your-server/Reports`

If you normally view reports in a browser, the URL you use is the **Report Portal URL**. The **Report Server URL** is usually the same server with `/ReportServer` instead of `/Reports`.

Then choose a folder and shared data source, name the report, and click **Create Report**.

Flank remembers the server URLs, folder and data source from the last report you created (per Windows user) and pre-fills them next time.

## Power BI Paginated Reports

Flank can publish a SQL query or stored procedure as a paginated report in a Power BI workspace. Paginated reports are SSRS reports hosted in Power BI, so you get the same basic tabular report as **SSRS Report**, without an SSRS server.

### Current support

- Power BI
  - Workspaces with a Power BI Pro or Premium Per User licence, or on Premium / Fabric capacity
  - Signing in with your work or school account, in the browser or with a code. Flank remembers the sign-in (per Windows user)
  - Reading the database through a **Gateway / Cloud Connection** you already have in Power BI. Flank picks the one for the database you're connected to and remembers your choice
  - Report parameters, with dropdowns from a stored procedure, a query, or values you type in
  - Publishing again with the same name in the same workspace replaces the report
- SQL
  - Selected SQL, or the entire query if nothing is selected
  - Parameters in a query (`WHERE league_id = @League_Id`), run as written or turned into a stored procedure
  - Stored procedures, from Object Explorer

### Not supported yet

- Power BI
  - Creating connections or gateways (create them in Power BI)
  - Viewers reading the database as themselves (single sign-on)
  - Report formatting and customization
- SQL
  - Multiple result sets

### Create a report

Right-click a query and choose **Make Self-Serve → Power BI (Paginated Report) (Beta)**, or right-click a stored procedure in Object Explorer. Sign in if Flank asks, then choose the workspace, name the report, pick the connection, and click **Publish**.

### One-time setup Flank doesn't do for you

Steps as of October 21, 2026; Microsoft moves these menus from time to time. If a step fails, Flank's error message says which one and what to check.

- **Approve Flank.** The first sign-in asks you to let Flank publish reports and read your connections. If your organization only lets admins approve apps, you'll see "Need admin approval": an admin (Cloud Application Administrator or Global Administrator) approves it in the **Microsoft Entra admin center → Enterprise applications → Admin consent requests**. The permissions are Power BI's Report.ReadWrite.All, Dataset.ReadWrite.All and Workspace.Read.All, and Fabric's Connection.Read.All.
- **A workspace you can publish to.** You need the **Contributor**, **Member** or **Admin** role on it (in Power BI: the workspace → **Manage access**).
- **A connection to your database.** In Power BI: **Settings (gear) → Manage connections and gateways**:
  - **Azure SQL Database:** a **cloud connection** (type SQL Server) to the server and database. The Azure SQL server's firewall must let Power BI in (Azure Portal → the SQL server → **Networking** → *Allow Azure services and resources to access this server*).
  - **SQL Server on your network:** an **on-premises data gateway** installed on a machine that can reach the server, and a gateway connection to the server and database.
  - On the connection's **Manage users**, add yourself (and anyone else who publishes with Flank) as a **User**.

## Uninstall

Remove Flank normally from **Windows Settings → Apps → Installed apps**.

Close SSMS before uninstalling.
