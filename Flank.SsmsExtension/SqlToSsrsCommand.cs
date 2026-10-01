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

            SsrsConnectionOptions connectionOptions;

            using (var dialog = new SsrsConnectionDialog())
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                connectionOptions = dialog.Options;
            }

            SsrsProgressDialog progressDialog = null;

            try
            {
                //
                // Step 1: Connect to SSRS and discover folders/data sources
                //

                var client = new SsrsClient(
                    connectionOptions.ReportServerUrl,
                    connectionOptions.ReportPortalUrl);

                progressDialog = new SsrsProgressDialog();
                progressDialog.SetStatus(
                    "Loading SSRS folders and data sources...");
                progressDialog.Show();
                progressDialog.Refresh();

                IReadOnlyList<string> folders;
                IReadOnlyList<string> dataSources;

                try
                {
                    var result = await Task.Run(() =>
                    {
                        client.TestConnection();

                        return new
                        {
                            Folders = client.GetFolders(),
                            DataSources = client.GetSharedDataSources()
                        };
                    });

                    folders = result.Folders;
                    dataSources = result.DataSources;
                }
                finally
                {
                    await ThreadHelper.JoinableTaskFactory
                        .SwitchToMainThreadAsync();

                    if (progressDialog != null)
                    {
                        progressDialog.Close();
                        progressDialog.Dispose();
                        progressDialog = null;
                    }
                }

                //
                // Step 2: Get report options
                //

                SsrsReportOptions reportOptions;

                using (var dialog = new SsrsTextDialog(
                    sql,
                    folders,
                    dataSources))
                {
                    if (dialog.ShowDialog() != DialogResult.OK)
                        return;

                    reportOptions = dialog.Options;
                }

                //
                // Step 3: Discover SQL schema
                //

                progressDialog = new SsrsProgressDialog();
                progressDialog.SetStatus("Analyzing query...");
                progressDialog.Show();
                progressDialog.Refresh();

                var connection =
                    SsmsQueryHelper.GetCurrentConnection();

                var columns =
                    SqlMetadataDiscovery.DiscoverQueryColumns(
                        reportOptions.Sql,
                        connection);

                // SSMS's connection is no longer needed.

                //
                // Step 4: Create report
                //

                progressDialog.SetStatus(
                    "Creating report in SSRS...");

                string reportUrl = await Task.Run(() =>
                {
                    return client.CreateReportFromText(
                        reportOptions.Sql,
                        columns,
                        reportOptions.DataSource,
                        reportOptions.Folder,
                        reportOptions.ReportName);
                });

                await ThreadHelper.JoinableTaskFactory
                    .SwitchToMainThreadAsync();

                progressDialog.Close();
                progressDialog.Dispose();
                progressDialog = null;

                //
                // Step 5: Success
                //

                using (var successDialog =
                    new SsrsSuccessDialog(
                        reportOptions.ReportName,
                        reportUrl))
                {
                    successDialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                await ThreadHelper.JoinableTaskFactory
                    .SwitchToMainThreadAsync();

                if (progressDialog != null)
                {
                    progressDialog.Close();
                    progressDialog.Dispose();
                    progressDialog = null;
                }

                SsmsUiHelper.ShowMessage(
                    "Flank couldn't create the SSRS report.\n\n" +
                    ex.Message,
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    "Flank Error");
            }
        }
    }
}
