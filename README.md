# Flank

Turn a SQL query into a refreshable Excel workbook directly from SSMS.

Instead of running a query, exporting the results to CSV, and sending someone a new file every time they need updated data, Flank lets you turn the query into an Excel workbook they can refresh themselves.

<img width="1600" height="900" alt="Flank SSMS Screenshot - Top of Menu2x" src="https://github.com/user-attachments/assets/af5316db-e739-45ee-a842-372c804066ed" />

## Requirements

Flank is currently an early experiment and has only been tested with:

- Windows
- SQL Server Management Studio 22
- Desktop Microsoft Excel with Power Query

The person refreshing the workbook also needs network access to the database and a way to authenticate to it (see below).

## Install

Download `Flank-SSMS-Setup.exe` from the latest GitHub release.

Close SSMS, run the installer, then reopen SSMS.

Right-click inside a SQL query and you should see **Make Self-Serve → Refreshable Excel**.

## How do I give someone access to refresh the workbook?

Flank doesn't put your database credentials in the workbook. When someone clicks **Refresh All**, Excel connects to the database and authenticates that user.

How you set that up depends on how your organization already manages database access:

**Users already connect to SQL Server with Windows/Active Directory authentication**

Give the recipient (or an AD group they're in) access to the database. Excel can connect using their Windows identity.

**You use Azure SQL with Microsoft Entra authentication**

Give the recipient (or preferably an Entra group they're in) access to the database. On first refresh, Excel will prompt them to authenticate with their Microsoft account.

**You use SQL Server authentication**

The recipient can enter their SQL Server username and password when Excel prompts them. Flank does not embed those credentials in the workbook.

**End users aren't allowed to connect to the database**

Flank's refreshable Excel output isn't currently a fit. The workbook connects directly from Excel to SQL Server; there is no Flank server sitting between the user and the database.

For least-privilege access, consider exposing the data through a stored procedure and granting users permission to execute that procedure rather than broad read access to the underlying tables.

## Uninstall

Remove Flank normally from **Windows Settings → Apps → Installed apps**.

Close SSMS before uninstalling.
