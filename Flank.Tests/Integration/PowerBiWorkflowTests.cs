using Flank.PowerBi;
using Flank.Ssrs;
using Flank.Tests.Support;

namespace Flank.Tests.Integration;

/// <summary>
/// The Power BI workflow after SSMS hands over the SQL or procedure: discover
/// its columns (and parameters), publish a paginated report with the database
/// embedded, connect it (a login on the report, or a shared connection), then
/// export it. Needs FLANK_TEST_SQL and FLANK_TEST_PBI.
/// </summary>
[Collection(IntegrationCollection.Name)]
[Trait("Category", "Integration")]
public class PowerBiWorkflowTests
{
    private readonly PowerBiTestWorkspace workspace;

    public PowerBiWorkflowTests(PowerBiTestWorkspace workspace)
    {
        this.workspace = workspace;
    }

    private static PowerBiDataAccess ReportLogin()
    {
        var (username, password) = PowerBiTestWorkspace.SqlLogin;
        return new PowerBiDataAccess.ReportLogin(username, password);
    }

    private async Task<PublishedReport> PublishQueryAsync(string sql, string reportName, PowerBiDataAccess access)
    {
        await workspace.EnsureReadyAsync();

        IReadOnlyList<ReportColumn> columns;
        using (var connection = TestSettings.OpenSqlConnection())
            columns = SqlMetadataDiscovery.DiscoverQueryColumns(sql, connection);

        string rdl = SsrsClient.BuildTextReportRdl(
            sql, columns, RdlDataSource.AzureSql(TestSettings.Server, TestSettings.Database));

        return await PowerBiPublisher.PublishAsync(
            workspace.Client, workspace.WorkspaceId, PowerBiTestWorkspace.ReportPrefix + reportName, rdl,
            TestSettings.Server, TestSettings.Database, access);
    }

    private async Task<PublishedReport> PublishSprocAsync(string procedure, string reportName)
    {
        await workspace.EnsureReadyAsync();

        IReadOnlyList<SprocParameter> parameters;
        IReadOnlyList<ReportColumn> columns;
        using (var connection = TestSettings.OpenSqlConnection())
        {
            int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(procedure, connection);
            parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(objectId, connection);
            columns = SqlMetadataDiscovery.DiscoverStoredProcedureColumns(objectId, procedure, connection);
        }

        string rdl = SsrsClient.BuildStoredProcedureReportRdl(
            procedure, parameters, columns, RdlDataSource.AzureSql(TestSettings.Server, TestSettings.Database));

        return await PowerBiPublisher.PublishAsync(
            workspace.Client, workspace.WorkspaceId, PowerBiTestWorkspace.ReportPrefix + reportName, rdl,
            TestSettings.Server, TestSettings.Database, ReportLogin());
    }

    [Fact]
    public async Task Query_report_with_a_login_on_the_report_exports_with_data()
    {
        var report = await PublishQueryAsync(
            "-- Leagues for the smoke test\r\nSELECT id, name\r\nFROM leagues\r\nWHERE name <> 'nobody'",
            "Query report",
            ReportLogin());

        // Check through the link, as the success dialog shows it.
        string output = Verifier.VerifyPowerBi(report.Url);
        Assert.Contains("Value_name", output);
    }

    [Fact]
    public async Task Publishing_again_replaces_the_report()
    {
        await PublishQueryAsync("SELECT id FROM leagues", "Replaced", ReportLogin());
        var second = await PublishQueryAsync("SELECT id, name FROM leagues", "Replaced", ReportLogin());

        var reports = await workspace.Client.GetReportsAsync(workspace.WorkspaceId);
        Assert.Single(reports, r => r.Name == PowerBiTestWorkspace.ReportPrefix + "Replaced");
        Assert.Contains("Value_name", Verifier.VerifyPowerBi(second.Url));
    }

    [Fact]
    public async Task Stored_procedure_report_with_parameters_exports_with_data()
    {
        var report = await PublishSprocAsync("dbo.sgm_ListTeamsByLeague", "Sproc report with parameters");

        string output = Verifier.VerifyPowerBi(report.Url, "League_Id=1");
        Assert.Contains("Value_team_name", output);
    }

    [Fact]
    public async Task Report_on_a_new_shared_connection_exports_and_the_connection_is_listed()
    {
        await workspace.EnsureReadyAsync();
        var (username, password) = PowerBiTestWorkspace.SqlLogin;
        string connectionName = "Flank test " + Guid.NewGuid().ToString("N").Substring(0, 8);

        var report = await PublishQueryAsync(
            "SELECT id, name FROM leagues",
            "Shared connection",
            new PowerBiDataAccess.NewConnection(connectionName, username, password));

        try
        {
            Assert.NotNull(report.ConnectionId);

            var dataSource = (await workspace.Client.GetReportDataSourcesAsync(workspace.WorkspaceId, report.ReportId)).Single();
            Assert.Equal(report.ConnectionId, dataSource.DatasourceId);

            // Listed as a cloud connection to the test database, as the dropdown shows it.
            var listed = (await workspace.Client.GetConnectionsAsync()).Single(c => c.Id == report.ConnectionId);
            Assert.Equal(connectionName, listed.DisplayName);
            Assert.Equal(PowerBiConnectionKind.Cloud, listed.Kind);
            Assert.True(listed.IsFor(TestSettings.Server, TestSettings.Database));

            Verifier.VerifyPowerBi(report.Url);

            // A second report reusing it, as when the user picks it in the dropdown.
            var reused = await PublishQueryAsync(
                "SELECT name FROM leagues",
                "Existing shared connection",
                new PowerBiDataAccess.ExistingConnection(listed));
            Verifier.VerifyPowerBi(reused.Url);
            await workspace.Client.DeleteReportAsync(workspace.WorkspaceId, reused.ReportId);
        }
        finally
        {
            await workspace.Client.DeleteReportAsync(workspace.WorkspaceId, report.ReportId);
            await workspace.Client.DeleteConnectionAsync(report.ConnectionId!);
        }
    }
}
