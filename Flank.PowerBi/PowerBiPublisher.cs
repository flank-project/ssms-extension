using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Flank.PowerBi
{
    /// <summary>How the published report reads the database.</summary>
    public abstract class PowerBiDataAccess
    {
        private PowerBiDataAccess()
        {
        }

        /// <summary>A SQL login saved on this report only.</summary>
        public sealed class ReportLogin : PowerBiDataAccess
        {
            public ReportLogin(string username, string password)
            {
                Username = username;
                Password = password;
            }

            public string Username { get; }

            public string Password { get; }
        }

        /// <summary>An existing shared connection (like an SSRS shared data source).</summary>
        public sealed class ExistingConnection : PowerBiDataAccess
        {
            public ExistingConnection(string connectionId)
            {
                ConnectionId = connectionId;
            }

            public string ConnectionId { get; }
        }

        /// <summary>A new shared connection with a SQL login, for this and later reports.</summary>
        public sealed class NewConnection : PowerBiDataAccess
        {
            public NewConnection(string displayName, string username, string password)
            {
                DisplayName = displayName;
                Username = username;
                Password = password;
            }

            public string DisplayName { get; }

            public string Username { get; }

            public string Password { get; }
        }
    }

    public sealed class PublishedReport
    {
        public PublishedReport(string workspaceId, string reportId, string? connectionId)
        {
            WorkspaceId = workspaceId;
            ReportId = reportId;
            ConnectionId = connectionId;
        }

        public string WorkspaceId { get; }

        public string ReportId { get; }

        /// <summary>The shared connection the report uses, if any (new or existing).</summary>
        public string? ConnectionId { get; }

        public string Url => PowerBiClient.GetReportUrl(WorkspaceId, ReportId);
    }

    /// <summary>
    /// Publishes a paginated report and connects it to the database: the
    /// steps after Flank has built the RDL, shared by SSMS and the tests.
    /// </summary>
    public static class PowerBiPublisher
    {
        public static async Task<PublishedReport> PublishAsync(
            PowerBiClient client,
            string workspaceId,
            string reportName,
            string rdl,
            string server,
            string database,
            PowerBiDataAccess access,
            IProgress<string>? progress = null,
            CancellationToken cancellationToken = default)
        {
            progress?.Report("Uploading the report to Power BI...");
            string reportId = await client.ImportReportAsync(workspaceId, reportName, rdl, cancellationToken).ConfigureAwait(false);

            var dataSource = (await client.GetReportDataSourcesAsync(workspaceId, reportId, cancellationToken).ConfigureAwait(false))
                .FirstOrDefault()
                ?? throw new PowerBiException("Power BI didn't find the report's data source.");

            progress?.Report("Connecting the report to the database...");

            switch (access)
            {
                case PowerBiDataAccess.ReportLogin login:
                    await client.SetSqlLoginAsync(dataSource, login.Username, login.Password, cancellationToken).ConfigureAwait(false);
                    return new PublishedReport(workspaceId, reportId, null);

                case PowerBiDataAccess.ExistingConnection existing:
                    await client.BindToConnectionAsync(workspaceId, reportId, dataSource, existing.ConnectionId, cancellationToken).ConfigureAwait(false);
                    return new PublishedReport(workspaceId, reportId, existing.ConnectionId);

                case PowerBiDataAccess.NewConnection create:
                    var connection = await client.CreateSqlConnectionAsync(
                        create.DisplayName, server, database, create.Username, create.Password, cancellationToken).ConfigureAwait(false);
                    await client.BindToConnectionAsync(workspaceId, reportId, dataSource, connection.Id, cancellationToken).ConfigureAwait(false);
                    return new PublishedReport(workspaceId, reportId, connection.Id);

                default:
                    throw new ArgumentOutOfRangeException(nameof(access));
            }
        }

        /// <summary>
        /// Power BI reaches Azure SQL directly; other servers need an on-premises
        /// data gateway, which Flank doesn't set up (yet).
        /// </summary>
        public static bool IsReachableFromPowerBi(string server)
        {
            string host = server.Trim();

            if (host.StartsWith("tcp:", StringComparison.OrdinalIgnoreCase))
                host = host.Substring(4);

            int comma = host.IndexOf(',');
            if (comma >= 0)
                host = host.Substring(0, comma);

            return host.EndsWith(".database.windows.net", StringComparison.OrdinalIgnoreCase);
        }
    }
}
