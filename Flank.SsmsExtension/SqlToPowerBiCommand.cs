using Flank.Ssrs;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Data.Common;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Make Self-Serve → Power BI Paginated Report in the query editor.
    /// </summary>
    internal static class SqlToPowerBiCommand
    {
        public const int CommandId = 0x0130;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory
                .SwitchToMainThreadAsync(package.DisposalToken);

            var commandService =
                await package.GetServiceAsync(typeof(IMenuCommandService))
                    as OleMenuCommandService;

            commandService?.AddCommand(new MenuCommand(
                (sender, e) => ThreadHelper.JoinableTaskFactory.RunAsync(ExecuteAsync),
                new CommandID(SqlToSsrsCommand.CommandSet, CommandId)));
        }

        private static async Task ExecuteAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            string sql;
            IDbConnection connection;

            try
            {
                sql = SsmsQueryHelper.GetCurrentSql();
                connection = SsmsQueryHelper.GetCurrentConnection();
            }
            catch (InvalidOperationException ex)
            {
                SsmsUiHelper.ShowMessage(ex.Message, OLEMSGICON.OLEMSGICON_INFO);
                return;
            }

            string server =
                SsmsSqlLogin.ActiveQueryWindowServer() ??
                (connection as DbConnection)?.DataSource ??
                "";

            await PowerBiWorkflow.RunAsync(
                server,
                connection.Database,
                SsmsSqlLogin.FromActiveQueryWindow(),
                sql,
                storedProcedure: null,
                defaultReportName: "SSMS Query",
                preparingStatus: "Analyzing query...",
                buildRdl: (options, dataSource) =>
                {
                    IReadOnlyList<ReportColumn> columns = null!;
                    SsmsQueryHelper.UseBorrowedConnection(connection, c =>
                        columns = SqlMetadataDiscovery.DiscoverQueryColumns(options.Sql!, c));

                    return SsrsClient.BuildTextReportRdl(options.Sql!, columns, dataSource);
                });
        }
    }
}
