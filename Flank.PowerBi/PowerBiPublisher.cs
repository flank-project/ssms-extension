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
            public ExistingConnection(PowerBiConnection connection)
            {
                Connection = connection;
            }

            public PowerBiConnection Connection { get; }
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
            string reportId;
            PowerBiReportDataSource dataSource;

            try
            {
                reportId = await client.ImportReportAsync(workspaceId, reportName, rdl, cancellationToken).ConfigureAwait(false);

                dataSource = (await client.GetReportDataSourcesAsync(workspaceId, reportId, cancellationToken).ConfigureAwait(false))
                    .FirstOrDefault()
                    ?? throw new PowerBiException("Power BI didn't find the report's data source.");
            }
            catch (Exception ex) when (!(ex is OperationCanceledException))
            {
                throw Tagged(ex, PowerBiStep.Upload);
            }

            progress?.Report("Connecting the report to the database...");

            try
            {
                return await ConnectAsync(client, workspaceId, reportId, dataSource, server, database, access, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex) when (!(ex is OperationCanceledException))
            {
                throw Tagged(ex, PowerBiStep.Connect);
            }
        }

        private static PowerBiException Tagged(Exception ex, PowerBiStep step) =>
            ex is PowerBiException { Step: not null } tagged
                ? tagged
                : new PowerBiException(ex.Message, ex) { Step = step };

        private static async Task<PublishedReport> ConnectAsync(
            PowerBiClient client,
            string workspaceId,
            string reportId,
            PowerBiReportDataSource dataSource,
            string server,
            string database,
            PowerBiDataAccess access,
            CancellationToken cancellationToken)
        {
            switch (access)
            {
                case PowerBiDataAccess.ReportLogin login:
                    await client.SetSqlLoginAsync(dataSource, login.Username, login.Password, cancellationToken).ConfigureAwait(false);
                    return new PublishedReport(workspaceId, reportId, null);

                case PowerBiDataAccess.ExistingConnection existing:
                    await client.BindToConnectionAsync(workspaceId, reportId, dataSource, existing.Connection, cancellationToken).ConfigureAwait(false);
                    return new PublishedReport(workspaceId, reportId, existing.Connection.Id);

                case PowerBiDataAccess.NewConnection create:
                    var connection = await client.CreateSqlConnectionAsync(
                        create.DisplayName, server, database, create.Username, create.Password, cancellationToken).ConfigureAwait(false);
                    await client.BindToConnectionAsync(workspaceId, reportId, dataSource, connection, cancellationToken).ConfigureAwait(false);
                    return new PublishedReport(workspaceId, reportId, connection.Id);

                default:
                    throw new ArgumentOutOfRangeException(nameof(access));
            }
        }
    }
}
