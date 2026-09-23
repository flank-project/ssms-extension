# Flank

Flank is an **SSMS extension** that turns a SQL query into a **refreshable Excel workbook.**

Right-click a query, choose `Make Self-Serve` → `Refreshable Excel`, and Flank creates an `.xlsx` file that you can send to an end user. Then, they can refresh the data themselves in Excel.

<img width="1600" height="900" alt="Flank SSMS Screenshot - Top of Menu2x" src="https://github.com/user-attachments/assets/af5316db-e739-45ee-a842-372c804066ed" />

## Current support

Flank current supports:

- **SSMS 22** on Windows
  - Selected SQL, or the entire query if nothing is selected
- **Desktop Excel** with Power Query
  - Power Query's "native SQL query"
  - Includes normal `SELECT` queries, CTEs, temp tables, variables, and stored procedures that return a result set
  - End user can "Refresh All" using their own database credentials

Not supported yet:

- Query parameters / user inputs
- Multiple result sets
- Power Query queries built from tables/views using Power Query transformations (query folding)
- Excel for the web
- End users who cannot connect directly to the database
- Older versions of SSMS / Excel

## Install

Download [Flank-SSMS-Setup-0.1.1.exe](https://github.com/flank-project/ssms-extension/releases/download/v0.1.1/Flank-SSMS-Setup-0.1.1.exe).

Close SSMS, run the installer, then reopen SSMS.

Right-click inside a SQL query and you should see **Make Self-Serve → Refreshable Excel**.

## How do I give an end user access?

The workbook connects directly to your database through Power Query, so the person refreshing it needs their own credentials. Flank does not put your credentials in the workbook.

If this is the first time you're giving an end user database access, there are a few common options.

### Azure SQL + Microsoft Entra

If you're using Azure SQL and the end user already has an account in your Microsoft Entra tenant, they can use that same account to authenticate from Excel.

#### 1. Check that Microsoft Entra authentication is enabled for your SQL server

In the Azure Portal:

1. Open the **SQL server** that contains your database (the logical server, not the individual database).
2. Under **Settings**, open **Microsoft Entra ID**.
3. Look for a **Microsoft Entra admin**.

If an admin is already listed, you're ready for the next step.

If not, click **Set admin**, select an Entra user or group, and click **Save**. This enables Microsoft Entra authentication for the logical server and establishes the Entra identity that can initially create other Entra users in SQL Server.

> This is a server-level setting, so you only need to configure it once for the Azure SQL logical server — not once per workbook or end user.

#### 2. Connect to the database using Microsoft Entra authentication

In SSMS, connect to the Azure SQL database using a Microsoft Entra authentication method rather than SQL Server authentication.

The account you connect with needs permission to create users in the database. If you're setting this up for the first time, connecting as the Microsoft Entra admin you configured above is the simplest option.

#### 3. Add the end user to the database

For an individual user:

```sql
CREATE USER [user@company.com] FROM EXTERNAL PROVIDER;
```

Or, if multiple people will refresh these workbooks, you can create an Entra group and add that group instead:

```sql
CREATE USER [Reporting Users] FROM EXTERNAL PROVIDER;
```

Then grant that user or group the permissions needed to run the query.

On their first refresh, Excel will prompt the end user to sign in with their Microsoft account. Excel then connects to Azure SQL as that user.
### SQL Server + Windows / Active Directory

If you're using SQL Server with Windows Authentication and your end users already have Windows/Active Directory accounts that SQL Server can recognize, you can use those identities instead.

Grant the individual Windows account or, more commonly, an AD group access to SQL Server and the database.

For example:

```sql
CREATE LOGIN [DOMAIN\Reporting Users] FROM WINDOWS;
CREATE USER [DOMAIN\Reporting Users] FOR LOGIN [DOMAIN\Reporting Users];
```

The end user can then select Windows authentication in Excel and connect using their existing Windows identity.

### SQL Server authentication

If neither of the above is available and you use SQL Server authentication (username + password), you can create a SQL login for the end user:

```sql
CREATE LOGIN report_user
WITH PASSWORD = '...';

CREATE USER report_user
FOR LOGIN report_user;
```

Excel will prompt the user for that username and password when they first refresh the workbook.

The tradeoff is that you're now managing another set of credentials. Users have another password to store, rotate, and potentially share, and identity/auditing is generally cleaner when you can use their existing Entra or Windows identity instead.

The credentials are handled by Excel and are not embedded in the workbook by Flank.

### What permissions does the user need?

The end user needs permission to execute whatever SQL is embedded in the workbook.

For example, if the workbook contains:

```sql
SELECT *
FROM dbo.vehicle_trips;
```

the user needs `SELECT` permission on `dbo.vehicle_trips` (either directly or through a role/group with that permission).

For a more constrained interface, put the query behind a stored procedure:

```sql
CREATE PROCEDURE dbo.GetVehicleTrips
AS
BEGIN
    SELECT *
    FROM dbo.vehicle_trips;
END;
```

Then grant the end user permission to execute only that procedure:

```sql
GRANT EXECUTE ON OBJECT::dbo.GetVehicleTrips
TO [Reporting Users];
```

The query embedded in the workbook can then simply be:

```sql
EXEC dbo.GetVehicleTrips;
```

This lets the user refresh the workbook without giving them direct `SELECT` permission on the underlying tables.

### What if end users aren't allowed to connect to the database?

Flank's refreshable Excel output isn't currently a fit for that environment.

Excel connects directly from the end user's machine to SQL Server. There is no Flank server or service account sitting between Excel and the database.

## Uninstall

Remove Flank normally from **Windows Settings → Apps → Installed apps**.

Close SSMS before uninstalling.
