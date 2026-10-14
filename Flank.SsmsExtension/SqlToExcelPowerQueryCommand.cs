using Flank.Excel;
using Flank.Ssrs;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    internal sealed class SqlToExcelPowerQueryCommand
    {
        public const int CommandId = 0x0100;

        public static readonly Guid CommandSet =
            new Guid("4aefd085-bfeb-4d3d-8199-201b0c86abc0");

        private readonly AsyncPackage package;

        private SqlToExcelPowerQueryCommand(
            AsyncPackage package,
            OleMenuCommandService commandService)
        {
            this.package = package
                ?? throw new ArgumentNullException(nameof(package));

            commandService = commandService
                ?? throw new ArgumentNullException(nameof(commandService));

            var commandId = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, commandId);

            commandService.AddCommand(menuItem);
        }

        public static SqlToExcelPowerQueryCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(
                package.DisposalToken);

            var commandService =
                await package.GetServiceAsync(typeof(IMenuCommandService))
                as OleMenuCommandService;

            Instance = new SqlToExcelPowerQueryCommand(
                package,
                commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await ExecuteAsync();
            });
        }

        private async Task ExecuteAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            string step = "Reading query";
            string server = null;
            string database = null;
            string sql;

            try
            {
                try
                {
                    sql = SsmsQueryHelper.GetCurrentSql();
                }
                catch (InvalidOperationException ex)
                {
                    SsmsUiHelper.ShowMessage(
                        ex.Message,
                        OLEMSGICON.OLEMSGICON_INFO);
                    return;
                }

                step = "Reading SQL Server connection";

                var scriptFactory = ServiceCache.ScriptFactory;

                if (scriptFactory == null)
                    throw new Exception(
                        "Could not access the SSMS connection service.");

                var connection =
                    scriptFactory.CurrentlyActiveWndConnectionInfo;

                var info = connection.UIConnectionInfo;

                if (info == null)
                    throw new Exception(
                        "No active SQL Server connection was found.");

                server = info.ServerName;
                database = info.AdvancedOptions["DATABASE"];

                if (string.IsNullOrWhiteSpace(server))
                    throw new Exception(
                        "Could not determine the SQL Server name.");

                if (string.IsNullOrWhiteSpace(database))
                    throw new Exception(
                        "Could not determine the database name.");
            }
            catch (Exception ex)
            {
                ExcelWorkflow.ShowFailure(
                    step,
                    ex,
                    server,
                    database,
                    null);
                return;
            }

            System.Data.IDbConnection queryConnection;
            bool isAzure = false;
            IReadOnlyList<string> tables = Array.Empty<string>();
            string tablesProblem = null;

            try
            {
                queryConnection = SsmsQueryHelper.GetCurrentConnection();

                SsmsQueryHelper.UseBorrowedConnection(queryConnection, connection =>
                {
                    isAzure = AccessDiscovery.IsAzureSqlDatabase(connection);

                    // The tables and views the query names, for granting SELECT.
                    try
                    {
                        tables = QueryObjects.Resolve(connection, QueryObjects.Find(sql));
                        if (tables.Count == 0)
                            tablesProblem = "none found in this database";
                    }
                    catch (Exception ex)
                    {
                        tablesProblem = ex.Message;
                    }
                });
            }
            catch (Exception ex)
            {
                ExcelWorkflow.ShowFailure(
                    "Reading SQL Server connection",
                    ex,
                    server,
                    database,
                    null);
                return;
            }

            // Who refreshes it, their access, then a run as them: the workbook
            // starts with that run's rows.
            WorkbookRequest request;
            string fileName;

            // The query window's connection is borrowed, never closed.
            using (var wizard = new ExcelWorkbookWizard(
                server,
                database,
                isAzure,
                procedure: null,
                procedureParameters: Array.Empty<SprocParameter>(),
                query: sql,
                queryTables: tables,
                queryTablesProblem: tablesProblem,
                action => SsmsQueryHelper.UseBorrowedConnection(queryConnection, action)))
            {
                if (wizard.ShowDialog() != System.Windows.Forms.DialogResult.OK || wizard.Request == null)
                    return;

                request = wizard.Request;
                fileName = wizard.DefaultFileName;
            }

            await ExcelWorkflow.RunAsync(request, fileName);
        }
    }
}
