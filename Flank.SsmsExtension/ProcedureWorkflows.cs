using Flank.Excel;
using Flank.Ssrs;
using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    /// <summary>A stored procedure to make self-serve, and how to reach its database.</summary>
    internal sealed class ProcedureSource
    {
        public ProcedureSource(string server, string database, string schema, string name, Func<IDbConnection> open)
        {
            Server = server;
            Database = database;
            Schema = schema;
            Name = name;
            Open = open;
        }

        public string Server { get; }

        public string Database { get; }

        public string Schema { get; }

        public string Name { get; }

        /// <summary>[schema].[name], quoted for T-SQL.</summary>
        public string QualifiedName => AccessScript.QuoteName(Schema) + "." + AccessScript.QuoteName(Name);

        /// <summary>Opens a connection to the procedure's database; the caller disposes it.</summary>
        public Func<IDbConnection> Open { get; }
    }

    /// <summary>
    /// Making a stored procedure self-serve: review its parameters, run it,
    /// then create the workbook or report. Shared by Object Explorer and by
    /// the query editor (after turning a query into a procedure).
    /// </summary>
    internal static class ProcedureWorkflows
    {
        public static async Task CreateExcelAsync(ProcedureSource sproc)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            string step = "Reading stored procedure";
            IReadOnlyList<SprocParameter> parameters;
            bool isAzure;

            try
            {
                using (var connection = sproc.Open())
                {
                    int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(
                        sproc.QualifiedName, connection);

                    parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(
                        objectId, connection);

                    isAzure = AccessDiscovery.IsAzureSqlDatabase(connection);
                }
            }
            catch (Exception ex)
            {
                ExcelWorkflow.ShowFailure(step, ex, sproc.Server, sproc.Database, null);
                return;
            }

            var settings = ReviewParameters(sproc, parameters, "workbook");
            if (settings == null)
                return;

            // Who refreshes it, their access, then a run as them: the workbook
            // starts with that run's values and rows.
            WorkbookRequest request;
            string fileName;

            using (var wizard = new ExcelWorkbookWizard(
                sproc.Server,
                sproc.Database,
                isAzure,
                sproc.QualifiedName,
                parameters,
                query: null,
                queryTables: Array.Empty<string>(),
                queryTablesProblem: null,
                action =>
                {
                    using (var connection = sproc.Open())
                    {
                        action(connection);
                    }
                },
                settings))
            {
                if (wizard.ShowDialog() != DialogResult.OK || wizard.Request == null)
                    return;

                request = wizard.Request;
                fileName = wizard.DefaultFileName;
            }

            await ExcelWorkflow.RunAsync(request, fileName);
        }

        public static async Task CreateSsrsReportAsync(ProcedureSource sproc)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            IReadOnlyList<SprocParameter> parameters;

            try
            {
                using (var connection = sproc.Open())
                {
                    int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(
                        sproc.QualifiedName, connection);

                    parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(
                        objectId, connection);
                }
            }
            catch (Exception ex)
            {
                SsmsUiHelper.ShowMessage(
                    "Flank couldn't read the stored procedure.\n\n" + ex.Message,
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    "Flank Error");
                return;
            }

            var settings = ReviewParameters(sproc, parameters, "report");
            if (settings == null)
                return;

            // Run it first to preview the results and get the report's columns.
            // The values entered aren't used as report defaults: those come
            // from the review page.
            ResultSet result;

            using (var dialog = new StoredProcedureRunDialog(
                sproc.QualifiedName,
                parameters,
                () => sproc.Open(),
                "Continue",
                "SSRS",
                settings))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                result = dialog.Results[0];
            }

            var columns = result.Columns
                .Select(name => new ReportColumn { Name = name })
                .ToList();

            await SsrsWorkflow.RunAsync(
                sql: null,
                storedProcedure: sproc.QualifiedName,
                defaultReportName: sproc.Name,
                analyzingStatus: "Preparing report...",
                prepare: reportOptions => client => client.CreateReportFromStoredProcedure(
                    sproc.QualifiedName,
                    parameters,
                    columns,
                    reportOptions.DataSource,
                    reportOptions.Folder,
                    reportOptions.ReportName,
                    settings: settings));
        }

        public static async Task CreatePowerBiReportAsync(ProcedureSource sproc)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            IReadOnlyList<SprocParameter> parameters;

            try
            {
                using (var connection = sproc.Open())
                {
                    int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(
                        sproc.QualifiedName, connection);

                    parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(
                        objectId, connection);
                }
            }
            catch (Exception ex)
            {
                SsmsUiHelper.ShowMessage(
                    "Flank couldn't read the stored procedure.\n\n" + ex.Message,
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    "Flank Error");
                return;
            }

            var settings = ReviewParameters(sproc, parameters, "report");
            if (settings == null)
                return;

            // As for SSRS: run it to preview the results and get the report's columns.
            ResultSet result;

            using (var dialog = new StoredProcedureRunDialog(
                sproc.QualifiedName,
                parameters,
                () => sproc.Open(),
                "Continue",
                "Power BI",
                settings))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                result = dialog.Results[0];
            }

            var columns = result.Columns
                .Select(name => new ReportColumn { Name = name })
                .ToList();

            await PowerBiWorkflow.RunAsync(
                sproc.Server,
                sproc.Database,
                sql: null,
                storedProcedure: sproc.QualifiedName,
                defaultReportName: sproc.Name,
                preparingStatus: "Preparing report...",
                buildRdl: (options, dataSource) => SsrsClient.BuildStoredProcedureReportRdl(
                    sproc.QualifiedName,
                    parameters,
                    columns,
                    dataSource,
                    settings: settings));
        }

        /// <summary>
        /// The parameter review page, for a procedure with input parameters.
        /// Null if the user cancels; empty (no page) without parameters.
        /// </summary>
        private static IReadOnlyList<ParameterSettings>? ReviewParameters(
            ProcedureSource sproc,
            IReadOnlyList<SprocParameter> parameters,
            string outputName) =>
            ReviewParameters(sproc.QualifiedName, parameters, outputName, sproc.Open);

        /// <summary>
        /// The parameter review page, for a procedure or query with input
        /// parameters. Null if the user cancels; empty (no page) without parameters.
        /// </summary>
        internal static IReadOnlyList<ParameterSettings>? ReviewParameters(
            string name,
            IReadOnlyList<SprocParameter> parameters,
            string outputName,
            Func<IDbConnection> openConnection)
        {
            var settings = ParameterSettings.FromProcedure(parameters);

            if (settings.Count == 0)
                return settings;

            using (var dialog = new ParameterReviewDialog(name, settings, outputName, openConnection))
            {
                return dialog.ShowDialog() == DialogResult.OK ? settings : null;
            }
        }
    }
}
