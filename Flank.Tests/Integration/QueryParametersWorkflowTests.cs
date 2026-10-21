using Flank.PowerBi;
using Flank.Ssrs;
using Flank.Tests.Support;

namespace Flank.Tests.Integration;

/// <summary>
/// A query from the query editor with @parameters: SQL Server suggests their
/// types, then the query runs, is described and becomes an SSRS or Power BI
/// report with them. Needs FLANK_TEST_SQL (and SSRS / FLANK_TEST_PBI).
/// </summary>
[Collection(IntegrationCollection.Name)]
[Trait("Category", "Integration")]
public class QueryParametersWorkflowTests
{
    private const string Sql =
        "-- Teams in a league\r\n" +
        "SELECT t.team_name, t.team_id\r\n" +
        "FROM dbo.sgm_teams AS t\r\n" +
        "WHERE t.league_id = @League_Id\r\n" +
        "  AND t.team_name LIKE @Name_Like";

    private readonly SsrsTestFolder folder;
    private readonly PowerBiTestWorkspace workspace;

    public QueryParametersWorkflowTests(SsrsTestFolder folder, PowerBiTestWorkspace workspace)
    {
        this.folder = folder;
        this.workspace = workspace;
    }

    private static IReadOnlyList<SprocParameter> Parameters()
    {
        using var connection = TestSettings.OpenSqlConnection();
        var suggested = QueryParameters.SuggestTypes(Sql, connection);

        return QueryParameters.Find(Sql)
            .Select(name => QueryParameters.Parameter(name, suggested[name])!)
            .ToList();
    }

    [Fact]
    public void Sql_server_suggests_the_parameter_types()
    {
        using var connection = TestSettings.OpenSqlConnection();

        var suggested = QueryParameters.SuggestTypes(Sql, connection);

        Assert.Equal("int", suggested["@League_Id"]);
        Assert.StartsWith("varchar(", suggested["@Name_Like"]);
    }

    [Fact]
    public void No_suggestions_for_a_query_sql_server_cannot_describe()
    {
        using var connection = TestSettings.OpenSqlConnection();

        Assert.Empty(QueryParameters.SuggestTypes("SELECT * FROM no_such_table WHERE id = @id", connection));
    }

    [Fact]
    public void The_query_runs_and_is_described_with_its_parameters()
    {
        var parameters = Parameters();

        using var connection = TestSettings.OpenSqlConnection();
        var results = StoredProcedureRunner.RunQuery(connection, Sql, parameters, new object?[] { 1, "%" }, 100);
        var columns = SqlMetadataDiscovery.DiscoverQueryColumns(Sql, connection, parameters);

        Assert.NotEmpty(results[0].Rows);
        Assert.Equal(new[] { "team_name", "team_id" }, results[0].Columns);
        Assert.Equal(new[] { "team_name", "team_id" }, columns.Select(c => c.Name));

        // Blank is NULL, so nothing matches.
        Assert.Empty(StoredProcedureRunner.RunQuery(connection, Sql, parameters, new object?[] { 1, null }, 100)[0].Rows);
    }

    [Fact]
    public void The_query_as_a_procedure_runs_with_its_parameters()
    {
        var parameters = Parameters();
        string name = "[dbo].[flank_test_query_" + Guid.NewGuid().ToString("N").Substring(0, 8) + "]";

        using var connection = TestSettings.OpenSqlConnection();
        using (var create = connection.CreateCommand())
        {
            create.CommandText = QueryParameters.CreateProcedureScript(name, parameters, Sql);
            create.ExecuteNonQuery();
        }

        try
        {
            var results = StoredProcedureRunner.Run(
                connection,
                name,
                new[] { new KeyValuePair<string, object?>("@League_Id", 1), new KeyValuePair<string, object?>("@Name_Like", "%") },
                100);

            Assert.NotEmpty(results[0].Rows);
        }
        finally
        {
            using var drop = connection.CreateCommand();
            drop.CommandText = "DROP PROCEDURE " + name;
            drop.ExecuteNonQuery();
        }
    }

    [Fact]
    public void Ssrs_query_report_with_parameters_renders_with_data()
    {
        var parameters = Parameters();
        var settings = ParameterSettings.FromProcedure(parameters);

        IReadOnlyList<ReportColumn> columns;
        using (var connection = TestSettings.OpenSqlConnection())
            columns = SqlMetadataDiscovery.DiscoverQueryColumns(Sql, connection, parameters);

        folder.EnsureClean();
        new SsrsClient(TestSettings.ReportServer, TestSettings.ReportPortal).CreateReportFromText(
            Sql, columns, parameters, settings, TestSettings.ReportDataSource, folder.Path, "Query with parameters");

        string output = Verifier.VerifySsrs(folder.Path + "/Query with parameters", "League_Id=1;Name_Like=%");
        Assert.Contains("Value_team_name", output);
    }

    [Fact]
    public async Task Power_bi_query_report_with_parameters_exports_with_data()
    {
        await workspace.EnsureReadyAsync();
        var parameters = Parameters();
        var settings = ParameterSettings.FromProcedure(parameters);

        IReadOnlyList<ReportColumn> columns;
        using (var connection = TestSettings.OpenSqlConnection())
            columns = SqlMetadataDiscovery.DiscoverQueryColumns(Sql, connection, parameters);

        string rdl = SsrsClient.BuildTextReportRdl(
            Sql, columns, RdlDataSource.AzureSql(TestSettings.Server, TestSettings.Database), parameters, settings);

        var (username, password) = PowerBiTestWorkspace.SqlLogin;
        var report = await PowerBiPublisher.PublishAsync(
            workspace.Client, workspace.WorkspaceId, PowerBiTestWorkspace.ReportPrefix + "Query with parameters", rdl,
            TestSettings.Server, TestSettings.Database, new PowerBiDataAccess.ReportLogin(username, password));

        string output = Verifier.VerifyPowerBi(report.Url, "League_Id=1;Name_Like=%");
        Assert.Contains("Value_team_name", output);
    }
}
