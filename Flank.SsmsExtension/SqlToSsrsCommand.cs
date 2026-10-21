using Flank.Ssrs;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SqlToSsrsCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0110;

        public static readonly Guid CommandSet =
            new Guid("4aefd085-bfeb-4d3d-8199-201b0c86abc0");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlToSsrsCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SqlToSsrsCommand(
            AsyncPackage package,
            OleMenuCommandService commandService)
        {
            this.package =
                package ?? throw new ArgumentNullException(nameof(package));

            commandService =
                commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID =
                new CommandID(CommandSet, CommandId);

            var menuItem =
                new MenuCommand(this.Execute, menuCommandID);

            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static SqlToSsrsCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory
                .SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService =
                await package.GetServiceAsync(typeof(IMenuCommandService))
                    as OleMenuCommandService;

            Instance =
                new SqlToSsrsCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// </summary>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await ExecuteAsync();
            });
        }

        private async Task ExecuteAsync()
        {
            await ThreadHelper.JoinableTaskFactory
                .SwitchToMainThreadAsync();

            string sql;
            System.Data.IDbConnection connection;

            try
            {
                sql = SsmsQueryHelper.GetCurrentSql();
                connection = SsmsQueryHelper.GetCurrentConnection();
            }
            catch (InvalidOperationException ex)
            {
                SsmsUiHelper.ShowMessage(
                    ex.Message,
                    OLEMSGICON.OLEMSGICON_INFO);

                return;
            }

            string server = SsmsQueryHelper.ActiveQueryWindowServer() ?? "";

            // As written or as a procedure, with any @parameters and their types.
            var setup = QueryWorkflows.SetUp(server, connection.Database, sql, connection, "report");
            if (setup == null)
                return;

            if (setup.Procedure != null)
            {
                await ProcedureWorkflows.CreateSsrsReportAsync(setup.Procedure);
                return;
            }

            ReviewedQuery reviewed = null;

            if (setup.Parameters.Count > 0)
            {
                reviewed = QueryWorkflows.ReviewAndRun(setup, connection, "report");
                if (reviewed == null)
                    return;
            }

            await SsrsWorkflow.RunAsync(
                setup.Sql,
                storedProcedure: null,
                defaultReportName: "SSMS Query",
                analyzingStatus: "Analyzing query...",
                prepare: reportOptions =>
                {
                    if (reviewed != null)
                    {
                        return client => client.CreateReportFromText(
                            setup.Sql,
                            reviewed.Columns,
                            setup.Parameters,
                            reviewed.Settings,
                            reportOptions.DataSource,
                            reportOptions.Folder,
                            reportOptions.ReportName);
                    }

                    IReadOnlyList<ReportColumn> columns = null;
                    SsmsQueryHelper.UseBorrowedConnection(connection, c =>
                        columns = SqlMetadataDiscovery.DiscoverQueryColumns(setup.Sql, c));

                    return client => client.CreateReportFromText(
                        setup.Sql,
                        columns,
                        reportOptions.DataSource,
                        reportOptions.Folder,
                        reportOptions.ReportName);
                });
        }
    }
}
