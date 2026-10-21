using Flank.PowerBi;
using Flank.Ssrs;
using Flank.Tests.Support;
using System.Data;
using System.Data.SqlClient;

namespace Flank.Tests.Integration;

/// <summary>
/// Reviewed parameter settings against a real procedure: defaults read from
/// its definition, blank values left out (so its default applies) or passed
/// as NULL, in the same call Excel, SSRS and Power BI run.
/// </summary>
[Collection(IntegrationCollection.Name)]
[Trait("Category", "Integration")]
public class ParameterSettingsWorkflowTests
{
    private readonly SsrsTestFolder folder;
    private readonly PowerBiTestWorkspace workspace;

    public ParameterSettingsWorkflowTests(SsrsTestFolder folder, PowerBiTestWorkspace workspace)
    {
        this.folder = folder;
        this.workspace = workspace;
    }

    private static IReadOnlyList<SprocParameter> Discover(out IReadOnlyList<ReportColumn> columns)
    {
        TestProcedures.EnsureCreated();

        using var connection = TestSettings.OpenSqlConnection();
        int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(TestProcedures.Parameters, connection);
        columns = SqlMetadataDiscovery.DiscoverStoredProcedureColumns(objectId, TestProcedures.Parameters, connection);
        return SqlMetadataDiscovery.DiscoverStoredProcedureParameters(objectId, connection);
    }

    [Fact]
    public void Defaults_are_discovered_from_the_procedure()
    {
        var parameters = Discover(out _).ToDictionary(p => p.Name);

        Assert.False(parameters["@League_Id"].HasDefault);
        Assert.True(parameters["@Name_Like"].HasDefault);
        Assert.Equal("%", parameters["@Name_Like"].DefaultValue);
        Assert.True(parameters["@Max_Id"].HasDefault);
        Assert.Null(parameters["@Max_Id"].DefaultValue);
        Assert.Equal("1", parameters["@Active"].DefaultValue);
        Assert.Equal("2000-01-01", parameters["@Since"].DefaultValue);
    }

    [Fact]
    public void The_call_leaves_out_blank_parameters_that_use_the_procedure_default()
    {
        var parameters = Discover(out _);
        var settings = ParameterSettings.FromProcedure(parameters);
        settings.Single(s => s.SqlName == "@Active").IfBlank = BlankValue.PassNull;

        string call = ProcedureCall.BuildExec("[dbo].[flank_test_params]", parameters, settings);

        // All blank except the league, as an empty Excel cell or report parameter would be.
        using var connection = TestSettings.OpenSqlConnection();
        using var command = connection.CreateCommand();
        command.CommandText = call;
        command.Parameters.Add(new SqlParameter("@League_Id", 1));
        command.Parameters.Add(new SqlParameter("@Name_Like", DBNull.Value));
        command.Parameters.Add(new SqlParameter("@Max_Id", DBNull.Value));
        command.Parameters.Add(new SqlParameter("@Active", DBNull.Value));
        command.Parameters.Add(new SqlParameter("@Since", DBNull.Value));

        var table = new DataTable();
        using (var reader = command.ExecuteReader())
            table.Load(reader);

        Assert.NotEmpty(table.Rows);
        var row = table.Rows[0];
        Assert.Equal("%", row["name_like"]);                              // the procedure's default
        Assert.Equal(new DateTime(2000, 1, 1), row["since"]);             // the procedure's default
        Assert.Equal(DBNull.Value, row["active"]);                        // NULL, as set, not the default 1
    }

    [Fact]
    public void Ssrs_report_with_reviewed_parameters_renders_with_only_the_required_one()
    {
        var parameters = Discover(out var columns);
        var settings = ParameterSettings.FromProcedure(parameters);
        settings[0].DisplayName = "League";

        folder.EnsureClean();
        new SsrsClient(TestSettings.ReportServer, TestSettings.ReportPortal).CreateReportFromStoredProcedure(
            "[dbo].[flank_test_params]", parameters, columns, TestSettings.ReportDataSource,
            folder.Path, "Reviewed parameters", settings: settings);

        // The optional parameters start as NULL, so the procedure's defaults apply.
        string output = Verifier.VerifySsrs(folder.Path + "/Reviewed parameters", "League_Id=1");
        Assert.Contains("Value_name_like", output);
    }

    [Fact]
    public async Task Power_bi_report_with_reviewed_parameters_exports_with_only_the_required_one()
    {
        await workspace.EnsureReadyAsync();
        var parameters = Discover(out var columns);
        var settings = ParameterSettings.FromProcedure(parameters);

        string rdl = SsrsClient.BuildStoredProcedureReportRdl(
            "[dbo].[flank_test_params]", parameters, columns,
            RdlDataSource.AzureSql(TestSettings.Server, TestSettings.Database), settings: settings);

        var (username, password) = PowerBiTestWorkspace.SqlLogin;
        var report = await PowerBiPublisher.PublishAsync(
            workspace.Client, workspace.WorkspaceId, PowerBiTestWorkspace.ReportPrefix + "Reviewed parameters", rdl,
            TestSettings.Server, TestSettings.Database, new PowerBiDataAccess.ReportLogin(username, password));

        Verifier.VerifyPowerBi(report.Url, "League_Id=1");
    }
}
