using Flank.PowerBi;
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
    /// Sign-in, dialogs and publishing for a Power BI paginated report, shared
    /// by the query editor and Object Explorer commands.
    /// </summary>
    internal static class PowerBiWorkflow
    {
        private const string Title = "Creating Power BI Report";

        /// <param name="server">The server as SSMS connected to it.</param>
        /// <param name="sql">The query, shown editable in the dialog. Null for a stored procedure.</param>
        /// <param name="storedProcedure">The stored procedure, shown read-only. Null for a query.</param>
        /// <param name="buildRdl">
        /// Runs on the UI thread once the options are chosen (e.g. to discover
        /// columns) and returns the report definition.
        /// </param>
        public static async Task RunAsync(
            string server,
            string database,
            string? sql,
            string? storedProcedure,
            string defaultReportName,
            string preparingStatus,
            Func<PowerBiReportOptions, RdlDataSource, string> buildRdl)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            server = NormalizeServer(server);

            SsrsProgressDialog? progressDialog = null;

            void ShowProgress(string status)
            {
                if (progressDialog == null)
                {
                    progressDialog = new SsrsProgressDialog(Title);
                    progressDialog.Show();
                }

                progressDialog.SetStatus(status);
                progressDialog.Refresh();
            }

            void CloseProgress()
            {
                progressDialog?.Close();
                progressDialog?.Dispose();
                progressDialog = null;
            }

            try
            {
                var signIn = await PowerBiSignIn.CreateAsync();
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                var token = await signIn.GetTokenSilentAsync();
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                if (token == null)
                {
                    using (var dialog = new PowerBiSignInDialog(
                        signIn,
                        "Flank publishes the report to Power BI as you. " +
                        "Sign in with your work or school account (the one you use for Power BI)."))
                    {
                        if (dialog.ShowDialog() != DialogResult.OK || dialog.Token == null)
                            return;

                        token = dialog.Token;
                    }
                }

                var client = new PowerBiClient(
                    async cancellationToken =>
                        (await signIn.GetTokenSilentAsync(null, cancellationToken))?.AccessToken
                        ?? throw new PowerBiException("Your Power BI sign-in has expired. Try again to sign in."),
                    async cancellationToken =>
                        (await signIn.GetTokenSilentAsync(PowerBiSignIn.ConnectionScopes, cancellationToken))?.AccessToken);

                ShowProgress("Loading your Power BI workspaces...");

                IReadOnlyList<PowerBiWorkspace> workspaces = await client.GetWorkspacesAsync();
                IReadOnlyList<PowerBiConnection>? connections = await TryGetConnectionsAsync(client);

                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                CloseProgress();

                var settings = PowerBiSettings.Load();
                PowerBiReportOptions options;

                using (var dialog = new PowerBiReportDialog(
                    token.Username,
                    server,
                    database,
                    workspaces,
                    connections,
                    async () =>
                    {
                        using (var allow = new PowerBiSignInDialog(
                            signIn,
                            "To list your Gateway / Cloud Connections, Flank needs one more permission " +
                            "(Power BI connections). Sign in again to grant it.",
                            PowerBiSignIn.ConnectionScopes))
                        {
                            if (allow.ShowDialog() != DialogResult.OK)
                                return null;
                        }

                        var allowed = await TryGetConnectionsAsync(client);
                        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                        if (allowed == null)
                        {
                            SsmsUiHelper.ShowMessage(
                                "Flank still can't list Power BI connections. Ask your admin to grant the " +
                                "Flank app the Power BI Connection.Read.All permission.",
                                OLEMSGICON.OLEMSGICON_WARNING);
                        }

                        return allowed;
                    },
                    sql,
                    storedProcedure,
                    defaultReportName,
                    settings))
                {
                    if (dialog.ShowDialog() != DialogResult.OK || dialog.Options == null)
                        return;

                    options = dialog.Options;
                }

                ShowProgress(preparingStatus);
                // The report reads the database the connection points at, through it.
                var connection = options.Connection;
                string rdl = buildRdl(options, RdlDataSource.ForPowerBiConnection(
                    connection.Server, connection.Database, connection.Kind == PowerBiConnectionKind.Gateway));

                // Created on the UI thread, so reports arrive there.
                var progress = new Progress<string>(status => progressDialog?.SetStatus(status));

                var published = await PowerBiPublisher.PublishAsync(
                    client,
                    options.Workspace.Id,
                    options.ReportName,
                    rdl,
                    connection.Server,
                    connection.Database,
                    new PowerBiDataAccess.ExistingConnection(connection),
                    progress);

                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                CloseProgress();

                settings.WorkspaceId = options.Workspace.Id;
                settings.ConnectionId = connection.Id;
                settings.Save();

                using (var success = new SsrsSuccessDialog(options.ReportName, published.Url, "Power BI Report Created"))
                {
                    success.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                CloseProgress();

                SsmsUiHelper.ShowMessage(
                    "Flank couldn't create the Power BI report.\n\n" + ex.Message,
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    "Flank Error");
            }
        }

        /// <summary>The Gateway / Cloud Connections, or null if Flank may not list them.</summary>
        private static async System.Threading.Tasks.Task<IReadOnlyList<PowerBiConnection>?> TryGetConnectionsAsync(PowerBiClient client)
        {
            try
            {
                if (!await client.CanUseConnectionsAsync())
                    return null;

                return await client.GetConnectionsAsync();
            }
            catch (Exception ex)
            {
                // The dialog then offers to sign in for the permission.
                System.Diagnostics.Debug.WriteLine($"Flank: could not list Power BI connections: {ex.Message}");
                return null;
            }
        }

        // "tcp:host,1433" -> "host": Power BI wants the plain host name.
        private static string NormalizeServer(string server)
        {
            string host = server.Trim();

            if (host.StartsWith("tcp:", StringComparison.OrdinalIgnoreCase))
                host = host.Substring(4);

            int comma = host.IndexOf(',');
            return comma >= 0 ? host.Substring(0, comma) : host;
        }
    }
}
