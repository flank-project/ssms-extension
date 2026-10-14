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

            var request = new WorkbookRequest(sql, server, database)
            {
                ColumnNames = TryDiscoverColumns(sql),
            };

            await ExcelWorkflow.RunAsync(request, "report.xlsx");
        }

        /// <summary>
        /// The query's result columns, so the workbook's table starts with them.
        /// Null if they can't be described without running the query (e.g. some
        /// temp-table batches); the first refresh then names them instead.
        /// </summary>
        private static List<string> TryDiscoverColumns(string sql)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var connection = SsmsQueryHelper.GetCurrentConnection();

                return SqlMetadataDiscovery.DiscoverQueryColumns(sql, connection)
                    .Select(c => c.Name)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Flank: could not discover query columns: {ex.Message}");
                return null;
            }
        }
    }
}
