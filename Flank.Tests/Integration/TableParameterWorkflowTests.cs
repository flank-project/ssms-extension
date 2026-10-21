using Flank.PowerBi;
using Flank.Ssrs;
using Flank.Tests.Support;

namespace Flank.Tests.Integration;

/// <summary>
/// Table-valued parameters against real procedures: their table types and
/// columns discovered, rows passed as JSON in the same call Excel, SSRS and
/// Power BI run, the end user's access to the type, and one-column tables
/// as multi-value report parameters.
/// </summary>
[Collection(IntegrationCollection.Name)]
[Trait("Category", "Integration")]
public class TableParameterWorkflowTests
{
    private readonly SsrsTestFolder folder;
    private readonly PowerBiTestWorkspace workspace;

    public TableParameterWorkflowTests(SsrsTestFolder folder, PowerBiTestWorkspace workspace)
    {
        this.folder = folder;
        this.workspace = workspace;
    }

    private static IReadOnlyList<SprocParameter> Discover(string procedure, out IReadOnlyList<ReportColumn> columns)
    {
        TestProcedures.EnsureCreated();

        using var connection = TestSettings.OpenSqlConnection();
        int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(procedure, connection);
        columns = SqlMetadataDiscovery.DiscoverStoredProcedureColumns(objectId, procedure, connection);
        return SqlMetadataDiscovery.DiscoverStoredProcedureParameters(objectId, connection);
    }

    [Fact]
    public void Table_types_and_their_fillable_columns_are_discovered()
    {
        var rows = Discover(TestProcedures.TableRows, out _).Single(p => p.Name == "@Rows");

        Assert.Equal("[dbo].[flank_test_rows]", rows.TableType);

        // The identity column is left out: SQL Server fills it in.
        Assert.Equal(new[] { "team_id", "label", "active", "since", "amount" }, rows.Columns.Select(c => c.Name));
        Assert.Equal(new[] { "int", "nvarchar", "bit", "date", "decimal" }, rows.Columns.Select(c => c.SqlType));
        Assert.False(rows.Columns[0].IsNullable);
        Assert.True(rows.Columns[1].IsNullable);
        Assert.Equal("nvarchar(50)", ProcedureCall.TypeDeclaration(rows.Columns[1]));
        Assert.Equal("decimal(9, 2)", ProcedureCall.TypeDeclaration(rows.Columns[4]));

        using var connection = TestSettings.OpenSqlConnection();
        Assert.True(SqlMetadataDiscovery.GetCompatibilityLevel(connection) >= 130);
    }

    [Fact]
    public void Rows_are_passed_as_json_and_arrive_typed()
    {
        var parameters = Discover(TestProcedures.TableRows, out _);
        var settings = ParameterSettings.FromProcedure(parameters);
        var rows = new List<IReadOnlyList<object?>>
        {
            new object?[] { 7, "It's \"quoted\"", true, new DateTime(2024, 1, 2), 12.5m },
            new object?[] { 8, null, false, null, null },
        };

        using var connection = TestSettings.OpenSqlConnection();
        var result = StoredProcedureRunner.RunCall(
            connection, "[dbo].[flank_test_table_rows]", parameters, settings, new object?[] { 3, rows }, 100).Single();

        Assert.Equal(2, result.Rows.Count);
        var first = result.Rows[0];
        Assert.Equal(3, first[0]);
        Assert.Equal(7, first[2]);
        Assert.Equal("It's \"quoted\"", first[3]);
        Assert.Equal(true, first[4]);
        Assert.Equal(new DateTime(2024, 1, 2), first[5]);
        Assert.Equal(12.5m, first[6]);
        Assert.Null(result.Rows[1][3]);
        Assert.Equal(false, result.Rows[1][4]);
    }

    [Fact]
    public void No_rows_is_an_empty_table()
    {
        var parameters = Discover(TestProcedures.TableRows, out _);

        using var connection = TestSettings.OpenSqlConnection();
        var result = StoredProcedureRunner.RunCall(
            connection, "[dbo].[flank_test_table_rows]", parameters, null, new object?[] { 3, null }, 100).Single();

        Assert.Empty(result.Rows);
    }

    [Fact]
    public void The_end_user_needs_execute_on_the_table_type_which_the_access_script_grants()
    {
        var parameters = Discover(TestProcedures.TableIds, out _);
        const string user = "flank_test_table_user";

        using var connection = TestSettings.OpenSqlConnection();

        void Execute(string sql)
        {
            using var command = connection.CreateCommand();
            command.CommandText = sql;
            command.ExecuteNonQuery();
        }

        Execute($"IF DATABASE_PRINCIPAL_ID(N'{user}') IS NULL CREATE USER [{user}] WITHOUT LOGIN;");
        Execute($"GRANT EXECUTE ON OBJECT::[dbo].[flank_test_table_ids] TO [{user}];");
        Execute($"REVOKE EXECUTE ON TYPE::[dbo].[flank_test_ids] FROM [{user}];");

        var values = new object?[] { new List<IReadOnlyList<object?>> { new object?[] { 1 } } };
        IReadOnlyList<ResultSet> Run() => AccessTester.RunAs(connection, user, () =>
            StoredProcedureRunner.RunCall(connection, "[dbo].[flank_test_table_ids]", parameters, null, values, 100));

        var denied = Assert.ThrowsAny<Exception>(Run);
        Assert.Contains("EXECUTE permission was denied", denied.Message);

        AccessScript.Run(connection, new AccessRequest
        {
            Name = user,
            ExistingUser = true,
            Scope = AccessScope.Procedure,
            Procedure = "[dbo].[flank_test_table_ids]",
            TableTypes = parameters.Where(p => p.IsTable).Select(p => p.TableType!).ToList(),
        });

        Assert.Single(Run());
    }

    [Fact]
    public void Ssrs_report_with_a_one_column_table_renders_the_values_given()
    {
        var parameters = Discover(TestProcedures.TableIds, out var columns);
        var settings = ParameterSettings.FromProcedure(parameters);

        folder.EnsureClean();
        new SsrsClient(TestSettings.ReportServer, TestSettings.ReportPortal).CreateReportFromStoredProcedure(
            "[dbo].[flank_test_table_ids]", parameters, columns, TestSettings.ReportDataSource,
            folder.Path, "Table of ids", settings: settings);

        string output = Verifier.VerifySsrs(folder.Path + "/Table of ids", "Ids=1;Ids=2", minRows: 2);
        Assert.Contains("2 rows", output);
    }

    [Fact]
    public async Task Power_bi_report_with_a_one_column_table_exports()
    {
        await workspace.EnsureReadyAsync();

        var parameters = Discover(TestProcedures.TableIds, out var columns);
        string rdl = SsrsClient.BuildStoredProcedureReportRdl(
            "[dbo].[flank_test_table_ids]", parameters, columns,
            RdlDataSource.AzureSql(TestSettings.Server, TestSettings.Database),
            settings: ParameterSettings.FromProcedure(parameters));

        var (username, password) = PowerBiTestWorkspace.SqlLogin;
        var report = await PowerBiPublisher.PublishAsync(
            workspace.Client, workspace.WorkspaceId, PowerBiTestWorkspace.ReportPrefix + "Table of ids", rdl,
            TestSettings.Server, TestSettings.Database, new PowerBiDataAccess.ReportLogin(username, password));

        Verifier.VerifyPowerBi(report.Url, "Ids=1;Ids=2", minRows: 2);
    }
}
