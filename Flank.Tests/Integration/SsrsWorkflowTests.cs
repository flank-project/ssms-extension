using Flank.Ssrs;
using Flank.Tests.Support;

namespace Flank.Tests.Integration;

/// <summary>
/// The SSRS workflow after SSMS hands over the SQL or procedure: discover
/// its columns (and parameters), deploy the report, then render it.
/// Needs FLANK_TEST_SQL for discovery and a local SSRS server.
/// </summary>
[Collection(IntegrationCollection.Name)]
[Trait("Category", "Integration")]
public class SsrsWorkflowTests
{
    private readonly SsrsTestFolder folder;
    private readonly SsrsClient client;

    public SsrsWorkflowTests(SsrsTestFolder folder)
    {
        this.folder = folder;
        client = new SsrsClient(TestSettings.ReportServer, TestSettings.ReportPortal);
    }

    private string CreateQueryReport(string sql, string reportName)
    {
        using var connection = TestSettings.OpenSqlConnection();
        var columns = SqlMetadataDiscovery.DiscoverQueryColumns(sql, connection);

        folder.EnsureClean();
        return client.CreateReportFromText(sql, columns, TestSettings.ReportDataSource, folder.Path, reportName);
    }

    private string CreateSprocReport(string procedure, string reportName)
    {
        using var connection = TestSettings.OpenSqlConnection();
        int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(procedure, connection);
        var parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(objectId, connection);
        var columns = SqlMetadataDiscovery.DiscoverStoredProcedureColumns(objectId, procedure, connection);

        folder.EnsureClean();
        return client.CreateReportFromStoredProcedure(
            procedure, parameters, columns, TestSettings.ReportDataSource, folder.Path, reportName);
    }

    [Fact]
    public void Query_report_renders_with_data()
    {
        string url = CreateQueryReport(
            "-- Leagues for the smoke test\r\nSELECT id, name\r\nFROM leagues\r\nWHERE name <> 'nobody'",
            "Query report");

        // Check through the portal link, as the success dialog shows it.
        string output = Verifier.VerifySsrs(url);
        Assert.Contains("Value_name", output);
    }

    [Fact]
    public void Query_report_with_colliding_column_names_renders()
    {
        CreateQueryReport(
            "SELECT id AS [league id], id AS league_id, name, name FROM leagues",
            "Colliding columns");

        Verifier.VerifySsrs(folder.Path + "/Colliding columns");
    }

    [Fact]
    public void Stored_procedure_report_renders_with_data()
    {
        CreateSprocReport("dbo.get_leagues", "Sproc report");

        string output = Verifier.VerifySsrs(folder.Path + "/Sproc report");
        Assert.Contains("Value_name", output);
    }

    [Fact]
    public void Stored_procedure_report_with_parameters_renders_with_data()
    {
        CreateSprocReport("dbo.sgm_ListTeamsByLeague", "Sproc report with parameters");

        string output = Verifier.VerifySsrs(folder.Path + "/Sproc report with parameters", "League_Id=1");
        Assert.Contains("Value_team_name", output);
    }
}
