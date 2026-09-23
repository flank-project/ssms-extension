# Flank

Flank is an **SSMS extension** that turns a SQL query into a **refreshable Excel workbook.**

Right-click a query, choose `Make Self-Serve` → `Refreshable Excel`, and Flank creates an `.xlsx` file that you can send to an end user. Then, they can refresh the data themselves in Excel.

<img width="1600" height="682" alt="Flank SSMS Screenshot - Top of Menu Thin" src="https://github.com/user-attachments/assets/8bc7f562-acf3-4a22-b4b6-1f8f2feb6be6" />

## Current support

Flank currently supports:

- **SSMS 22** on Windows
  - Selected SQL, or the entire query if nothing is selected
- **Desktop Excel** with Power Query
  - End users can "Refresh All" using their own database credentials
- **Native SQL queries**
  - Uses Power Query's "native SQL query" support
  - Includes `SELECT` queries, CTEs, temp tables, variables, and stored procedures that return a result set

Not supported yet:

- SQL
  - Query parameters / user inputs
  - Multiple result sets
- Excel
  - Excel for the web
- SSMS
  - Versions earlier than SSMS 22

## Install

Download [Flank-SSMS-Setup-0.1.2.exe](https://github.com/flank-project/ssms-extension/releases/download/v0.1.2/Flank-SSMS-Setup-0.1.2.exe).

Close SSMS, run the installer, then reopen SSMS.

Right-click inside a SQL query and you should see **Make Self-Serve → Refreshable Excel**.

## How do I give an end user access?

The workbook connects directly to your database through Power Query, so the person refreshing it needs their own credentials. Flank does not put your credentials in the workbook.

If this is the first time you're giving an end user database access, there are a few common options.

### Azure SQL + Microsoft Entra

If you're using Azure SQL and the end user already has an account in your Microsoft Entra tenant, they can use that identity to refresh the workbook.

#### 1. Check that your Azure SQL server has a Microsoft Entra admin

In the Azure Portal:

1. Open the **SQL server** that contains your database (the logical server, not the individual database).
2. Under **Settings**, open **Microsoft Entra ID**.
3. Look for a **Microsoft Entra admin**.

If an admin is already listed, continue to the next step.

If not, click **Set admin**, select an Entra user or group, and click **Save**.

This is a server-level setting, so you only need to configure it once for the Azure SQL logical server.

#### 2. Connect to the database as the Microsoft Entra admin

In SSMS, connect to the database using **Microsoft Entra MFA** authentication and the Entra admin account from the previous step.

#### 3. Create a database user for the end user

Run:

```sql
CREATE USER [user@company.com] FROM EXTERNAL PROVIDER;
```

`user@company.com` should be the user's **User Principal Name (UPN)** in Microsoft Entra. This often looks like their email address, but the two can be different.

You can find the user's UPN in the Azure Portal under **Microsoft Entra ID → Users → [user] → User principal name**.

#### 4. Give the user access to the data

For the simplest setup, add the user to the built-in `db_datareader` role:

```sql
ALTER ROLE db_datareader ADD MEMBER [user@company.com];
```

This allows the user to read all user tables and views in that database.

That's intentionally broad. It's a convenient way to get your first workbook working, but you can narrow the user's permissions later.

For example, you can grant `SELECT` on only the tables/views the workbook needs, or put the query behind a stored procedure and grant the user `EXECUTE` permission on that procedure.

#### 5. Send the workbook

Send the generated `.xlsx` file to the end user. They do **not** need Flank installed.

When they click **Data → Refresh All** for the first time, Excel will ask them to authenticate to the database. Choose the Microsoft/Entra authentication option and sign in using the same Entra account you added above.

After authentication, Excel will run the workbook's query as that user and load the results.

### SQL Server + Windows / Active Directory

If you're using SQL Server with Windows Authentication and the end user already has a Windows/Active Directory account that SQL Server can recognize, they can use that identity to refresh the workbook.

This assumes the end user's computer can reach the SQL Server — for example, because they're on the corporate network or connected through a VPN.

#### 1. Check the end user's Windows identity

The user will normally have an Active Directory identity that looks something like:

```text
COMPANY\jsmith
```

This is the identity SQL Server will use when the user connects with Windows Authentication.

If you're not sure of the username, the end user can open Command Prompt and run:

```cmd
whoami
```

#### 2. Create a SQL Server login for the end user

In SSMS, connect to the SQL Server as an administrator and run:

```sql
CREATE LOGIN [COMPANY\jsmith] FROM WINDOWS;
```

This allows that Windows identity to authenticate to SQL Server.

#### 3. Create a user in the database

Switch to the database containing the data:

```sql
USE MyDatabase;
GO

CREATE USER [COMPANY\jsmith]
FOR LOGIN [COMPANY\jsmith];
```

The login gives the user access to the SQL Server. The database user gives that login an identity inside this particular database.

#### 4. Give the user access to the data

For the simplest setup, add the user to the built-in `db_datareader` role:

```sql
ALTER ROLE db_datareader ADD MEMBER [COMPANY\jsmith];
```

This allows the user to read all user tables and views in that database.

That's intentionally broad. It's a convenient way to get your first workbook working, but you can narrow the user's permissions later.

For example, you can grant `SELECT` on only the tables/views the workbook needs, or put the query behind a stored procedure and grant the user `EXECUTE` permission on that procedure.

#### 5. Send the workbook

Send the generated `.xlsx` file to the end user. They do **not** need Flank installed.

When they click **Data → Refresh All** for the first time, Excel will ask them to authenticate to the database. Choose **Windows** authentication.

Excel will connect to SQL Server using their Windows identity, run the workbook's query as that user, and load the results.

### SQL Server authentication

If you don't have Microsoft Entra or Windows/Active Directory authentication available and your SQL Server accepts SQL Server authentication (username + password), you can create a SQL login for the end user.

This assumes the end user's computer can reach the SQL Server — for example, because they're on the corporate network, connected through a VPN, or the server is otherwise reachable from their machine.

#### 1. Check that SQL Server authentication is enabled

In SSMS:

1. Right-click the SQL Server in **Object Explorer** and select **Properties**.
2. Open **Security**.
3. Under **Server authentication**, check that **SQL Server and Windows Authentication mode** is selected.

If you change this setting, SQL Server needs to be restarted before the change takes effect.

#### 2. Create a SQL Server login for the end user

Connect to SQL Server as an administrator and run:

```sql
CREATE LOGIN report_user
WITH PASSWORD = 'use-a-strong-password-here';
```

This creates a username and password that the end user can use to authenticate to SQL Server.

#### 3. Create a user in the database

Switch to the database containing the data:

```sql
USE MyDatabase;
GO

CREATE USER report_user
FOR LOGIN report_user;
```

The login gives the user access to the SQL Server. The database user gives that login an identity inside this particular database.

#### 4. Give the user access to the data

For the simplest setup, add the user to the built-in `db_datareader` role:

```sql
ALTER ROLE db_datareader ADD MEMBER report_user;
```

This allows the user to read all user tables and views in that database.

That's intentionally broad. It's a convenient way to get your first workbook working, but you can narrow the user's permissions later.

For example, you can grant `SELECT` on only the tables/views the workbook needs, or put the query behind a stored procedure and grant the user `EXECUTE` permission on that procedure.

#### 5. Send the workbook

Send the generated `.xlsx` file to the end user. They do **not** need Flank installed.

When they click **Data → Refresh All** for the first time, Excel will ask them to authenticate to the database. Choose **Database** authentication and enter the SQL Server username and password you created above.

Excel will connect to SQL Server using those credentials, run the workbook's query as that user, and load the results.

### Tradeoffs

SQL Server authentication works, but it means creating and managing a separate database password for the end user.

If Microsoft Entra or Windows authentication is available, those options are generally easier to manage because the user can authenticate with an identity they already have.

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
