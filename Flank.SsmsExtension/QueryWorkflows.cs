using Flank.Ssrs;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    /// <summary>A parameterized query, reviewed and run: how its parameters are set up, and its columns.</summary>
    internal sealed class ReviewedQuery
    {
        public ReviewedQuery(IReadOnlyList<ParameterSettings> settings, IReadOnlyList<ReportColumn> columns)
        {
            Settings = settings;
            Columns = columns;
        }

        public IReadOnlyList<ParameterSettings> Settings { get; }

        public IReadOnlyList<ReportColumn> Columns { get; }
    }

    /// <summary>
    /// The query editor's first steps, shared by Excel, SSRS and Power BI:
    /// the query setup (as written, or as a stored procedure, with
    /// parameters), then for parameterized SQL the review page and a run.
    /// </summary>
    internal static class QueryWorkflows
    {
        public const string QueryName = "the query";

        /// <summary>The query setup; null if the user cancels.</summary>
        /// <param name="borrowed">The query window's connection, lent (never closed).</param>
        public static QuerySetup? SetUp(string server, string database, string sql, IDbConnection borrowed, string outputName)
        {
            using (var dialog = new QuerySetupDialog(server, database, sql, outputName, () => new BorrowedConnection(borrowed)))
            {
                return dialog.ShowDialog() == DialogResult.OK ? dialog.Setup : null;
            }
        }

        /// <summary>
        /// For SQL with parameters (to a report): the review page, then a run
        /// to preview the results and get the columns. Null if the user cancels.
        /// </summary>
        public static ReviewedQuery? ReviewAndRun(QuerySetup setup, IDbConnection borrowed, string outputName)
        {
            Func<IDbConnection> open = () => new BorrowedConnection(borrowed);

            var settings = ProcedureWorkflows.ReviewParameters(QueryName, setup.Parameters, outputName, open);
            if (settings == null)
                return null;

            // Back returns to the review page, then runs again with what it says.
            while (true)
            {
                using (var dialog = new StoredProcedureRunDialog(
                    QueryName,
                    setup.Parameters,
                    open,
                    "Continue",
                    outputName,
                    settings,
                    values =>
                    {
                        using (var connection = open())
                        {
                            return StoredProcedureRunner.RunQuery(connection, setup.Sql, setup.Parameters, values, PreviewPanel.MaxRows);
                        }
                    },
                    canGoBack: settings.Count > 0))
                {
                    var answer = dialog.ShowDialog();

                    if (answer == DialogResult.Retry)
                    {
                        if (!ProcedureWorkflows.Review(QueryName, settings, outputName, open))
                            return null;
                        continue;
                    }

                    if (answer != DialogResult.OK)
                        return null;

                    var columns = dialog.Results[0].Columns
                        .Select(name => new ReportColumn { Name = name })
                        .ToList();

                    return new ReviewedQuery(settings, columns);
                }
            }
        }
    }
}
