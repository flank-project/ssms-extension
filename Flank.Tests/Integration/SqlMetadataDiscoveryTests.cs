using Flank.Ssrs;
using Flank.Tests.Support;

namespace Flank.Tests.Integration;

/// <summary>Column and parameter discovery against the test database. Needs FLANK_TEST_SQL.</summary>
[Collection(IntegrationCollection.Name)]
[Trait("Category", "Integration")]
public class SqlMetadataDiscoveryTests
{
    [Fact]
    public void Discovers_query_columns_without_running_the_query()
    {
        using var connection = TestSettings.OpenSqlConnection();

        var columns = SqlMetadataDiscovery.DiscoverQueryColumns("SELECT id, name AS league FROM leagues", connection);

        Assert.Equal(new[] { "id", "league" }, columns.Select(c => c.Name));
    }

    [Fact]
    public void Rejects_a_query_with_no_result_set()
    {
        using var connection = TestSettings.OpenSqlConnection();

        Assert.Throws<InvalidOperationException>(() =>
            SqlMetadataDiscovery.DiscoverQueryColumns("DECLARE @x int", connection));
    }

    [Fact]
    public void Discovers_procedure_parameters_and_columns()
    {
        using var connection = TestSettings.OpenSqlConnection();

        int id = SqlMetadataDiscovery.GetStoredProcedureObjectId("dbo.sgm_ListTeamsByLeague", connection);
        var parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(id, connection);
        var columns = SqlMetadataDiscovery.DiscoverStoredProcedureColumns(id, "dbo.sgm_ListTeamsByLeague", connection);

        var parameter = Assert.Single(parameters);
        Assert.Equal("@League_Id", parameter.Name);
        Assert.Equal("int", parameter.SqlType);
        Assert.Equal(new[] { "team_name", "team_id", "league_id" }, columns.Select(c => c.Name));
    }

    [Fact]
    public void Marks_output_parameters()
    {
        using var connection = TestSettings.OpenSqlConnection();

        int id = SqlMetadataDiscovery.GetStoredProcedureObjectId("dbo.usp_OnlyOutputParameters", connection);
        var parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(id, connection);

        Assert.Equal(
            new[] { ("@InputValue", false), ("@OutputIntValue", true), ("@OutputVarcharValue", true) },
            parameters.Select(p => (p.Name, p.IsOutput)));
    }

    [Fact]
    public void Reports_a_missing_procedure()
    {
        using var connection = TestSettings.OpenSqlConnection();

        Assert.Throws<InvalidOperationException>(() =>
            SqlMetadataDiscovery.GetStoredProcedureObjectId("dbo.does_not_exist_flank", connection));
    }
}
