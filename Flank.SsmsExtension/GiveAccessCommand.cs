using Flank.Ssrs;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Query editor: Make Self-Serve > Give Someone Access. Grants SELECT on
    /// the tables the query reads (or db_datareader).
    /// </summary>
    internal sealed class GiveAccessCommand
    {
        public const int CommandId = 0x0120;

        public static readonly Guid CommandSet =
            new Guid("4aefd085-bfeb-4d3d-8199-201b0c86abc0");

        private GiveAccessCommand(OleMenuCommandService commandService)
        {
            commandService.AddCommand(
                new MenuCommand(Execute, new CommandID(CommandSet, CommandId)));
        }

        public static GiveAccessCommand? Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(
                package.DisposalToken);

            if (await package.GetServiceAsync(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                Instance = new GiveAccessCommand(commandService);
            }
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            System.Data.IDbConnection connection;
            string sql;
            string server;

            try
            {
                sql = SsmsQueryHelper.GetCurrentSql();
                connection = SsmsQueryHelper.GetCurrentConnection();
                server = ServiceCache.ScriptFactory?.CurrentlyActiveWndConnectionInfo?.UIConnectionInfo?.ServerName
                    ?? (connection as System.Data.Common.DbConnection)?.DataSource
                    ?? "";
            }
            catch (InvalidOperationException ex)
            {
                SsmsUiHelper.ShowMessage(ex.Message, OLEMSGICON.OLEMSGICON_INFO);
                return;
            }

            bool isAzure;
            IReadOnlyList<string> tables;

            try
            {
                isAzure = AccessDiscovery.IsAzureSqlDatabase(connection);

                // The narrowest grant: the tables the query's columns come from.
                try
                {
                    tables = AccessDiscovery.QueryTables(connection, sql);
                }
                catch (Exception)
                {
                    tables = Array.Empty<string>();
                }
            }
            catch (Exception ex)
            {
                SsmsUiHelper.ShowMessage(
                    "Flank couldn't read the database.\n\n" + ex.Message,
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    "Flank Error");
                return;
            }

            // The query window's connection is borrowed, never closed.
            using (var dialog = new GiveAccessDialog(
                server,
                connection.Database,
                isAzure,
                procedure: null,
                tables,
                action => action(connection)))
            {
                dialog.ShowDialog();
            }
        }
    }
}
