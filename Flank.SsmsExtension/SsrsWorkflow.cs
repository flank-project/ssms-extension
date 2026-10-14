using Flank.Ssrs;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// The SSRS dialogs and report creation shared by the query editor and
    /// Object Explorer commands.
    /// </summary>
    internal static class SsrsWorkflow
    {
        /// <param name="sql">The query, shown editable in the report dialog. Null for a stored procedure.</param>
        /// <param name="storedProcedure">The stored procedure, shown read-only. Null for a query.</param>
        /// <param name="prepare">
        /// Runs on the UI thread once the report options are chosen (e.g. to
        /// discover columns). Returns the step that creates the report, which
        /// runs in the background and returns the report URL.
        /// </param>
        public static async Task RunAsync(
            string? sql,
            string? storedProcedure,
            string defaultReportName,
            string analyzingStatus,
            Func<SsrsReportOptions, Func<SsrsClient, string>> prepare)
        {
            await ThreadHelper.JoinableTaskFactory
                .SwitchToMainThreadAsync();

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
                    dataSources,
                    defaultReportName,
                    storedProcedure))
                {
                    if (dialog.ShowDialog() != DialogResult.OK)
                        return;

                    reportOptions = dialog.Options;
                }

                //
                // Step 3: Discover SQL schema
                //

                progressDialog = new SsrsProgressDialog();
                progressDialog.SetStatus(analyzingStatus);
                progressDialog.Show();
                progressDialog.Refresh();

                var deploy = prepare(reportOptions);

                //
                // Step 4: Create report
                //

                progressDialog.SetStatus(
                    "Creating report in SSRS...");

                string reportUrl = await Task.Run(() => deploy(client));

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
