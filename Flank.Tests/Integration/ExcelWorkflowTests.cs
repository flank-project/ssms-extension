using Flank.Excel;
using Flank.Tests.Support;

namespace Flank.Tests.Integration;

/// <summary>
/// The Excel workflow after SSMS hands over the SQL: generate the workbook,
/// then have Excel refresh it against the real database. Uses Excel's saved
/// credentials for the database, so FLANK_TEST_SQL isn't needed.
/// </summary>
[Collection(IntegrationCollection.Name)]
[Trait("Category", "Integration")]
public class ExcelWorkflowTests : IDisposable
{
    private readonly string outputPath =
        Path.Combine(Path.GetTempPath(), $"flank-integration-{Guid.NewGuid():N}.xlsx");

    public void Dispose()
    {
        File.Delete(outputPath);
    }

    [Fact]
    public void Query_workbook_refreshes_with_data()
    {
        // Shaped like a real query window: comments, line breaks, quoted names.
        string sql =
            "-- Leagues for the smoke test\r\n" +
            "SELECT id,\r\n" +
            "       name AS \"League Name\"\r\n" +
            "FROM leagues\r\n" +
            "WHERE name <> 'nobody'";

        ExcelExporter.Generate(outputPath, sql, TestSettings.Server, TestSettings.Database);

        string output = Verifier.VerifyExcel(outputPath);
        Assert.Contains("League Name", output);
    }

    [Fact]
    public void Stored_procedure_workbook_refreshes_with_data()
    {
        ExcelExporter.Generate(
            outputPath,
            "EXEC dbo.sgm_ListTeamsByLeague @League_Id = 1",
            TestSettings.Server,
            TestSettings.Database);

        string output = Verifier.VerifyExcel(outputPath);
        Assert.Contains("team_name", output);
    }
}
