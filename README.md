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

## How do I give an end user access?

The Excel workbook connects directly to your database, so the person refreshing it needs their own way to authenticate. Flank does not put your credentials in the workbook.

If your end users don't already have database access, here's the usual setup:

**Azure SQL / Microsoft Entra**

This is usually the simplest option if your organization already uses Microsoft Entra.

Add the user to the database with their existing work Microsoft account, or create an Entra group for users who should be able to refresh these workbooks.

On their first refresh, Excel will ask them to sign in with that account.

**SQL Server with Windows / Active Directory**

Grant access to the user's Windows/AD account, or preferably to an AD group they're a member of.

Excel can then connect using their Windows identity.

**SQL Server authentication**

Create a SQL login for the user and give it the necessary database permissions.

On their first refresh, Excel will ask for the SQL username and password. Those credentials are stored by Excel on the user's machine, not embedded by Flank in the workbook.

**Your organization doesn't allow end users to connect directly to databases**

Flank's refreshable Excel output isn't currently a fit. Excel connects directly to the database; there is no Flank server or service account sitting between the user and SQL Server.

### What permissions should I give them?

Avoid giving users broad access just to refresh a workbook.

For repeatable reports, a good pattern is to put the query behind a stored procedure and grant the user (or group) permission to execute that procedure rather than read access to all of the underlying tables.

## Uninstall

Remove Flank normally from **Windows Settings → Apps → Installed apps**.

Close SSMS before uninstalling.
