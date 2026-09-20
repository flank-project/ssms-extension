# Flank

Turn a SQL query into a refreshable Excel workbook directly from SSMS.

Instead of running a query, exporting the results to CSV, and sending someone a new file every time they need updated data, Flank lets you turn the query into an Excel workbook they can refresh themselves.

<img width="1600" height="900" alt="Flank SSMS Screenshot - Top of Menu2x" src="https://github.com/user-attachments/assets/af5316db-e739-45ee-a842-372c804066ed" />


## How it works

In SSMS:

1. Write or open a query.
2. Right-click in the query editor.
3. Choose **Make Self-Serve → Refreshable Excel**.
4. Save the workbook.

Flank creates an `.xlsx` file with the query embedded as a Power Query connection.

Open the workbook and click **Data → Refresh All** to fetch the latest data.

## Install

**Requirements**

- Windows
- SQL Server Management Studio 22
- Microsoft Excel with Power Query

Download `Flank-SSMS-Setup.exe` from the latest GitHub release.

Close SSMS, run the installer, then reopen SSMS.

That's it. There should now be a **Make Self-Serve** option when you right-click inside the SQL query editor.

## Authentication

Flank does **not** put your database credentials in the workbook.

The workbook contains the SQL query and connection information. When another user refreshes it, Excel authenticates that user to SQL Server using their own credentials.

For Azure SQL with Microsoft Entra authentication, this means you can give users access to the underlying database objects — ideally through a stored procedure or another narrowly permissioned interface — without sharing your own credentials.

## Why?

A pretty common SQL Server workflow looks like this:

    Someone asks for data
            ↓
    Write a query in SSMS
            ↓
    Save Results As...
            ↓
    Send CSV
            ↓
    "Can you send me updated numbers?"

For frequently reused data, you might eventually build an SSRS or Power BI report.

Flank is for the space in between.

If a query already answers the question, turning it into something another person can refresh should take seconds.

## Current status

Flank is an early experiment.

Right now it does one thing:

**SSMS query → refreshable Excel workbook**

Currently tested with SSMS 22 and Azure SQL using Microsoft Entra authentication.

If you try it and something breaks, please open an issue and include a screenshot of the Flank error message.

## Uninstall

Flank can be removed normally from **Windows Settings → Apps → Installed apps**.

Close SSMS before uninstalling.
