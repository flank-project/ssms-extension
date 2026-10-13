using Flank.Excel;
using Flank.Tests.Support;
using System.IO.Compression;

namespace Flank.Tests.Unit;

public class ExcelExporterTests : IDisposable
{
    private readonly string outputPath =
        Path.Combine(Path.GetTempPath(), $"flank-unit-{Guid.NewGuid():N}.xlsx");

    public void Dispose()
    {
        File.Delete(outputPath);
    }

    private string Generate(string sql, string server = "my-server", string database = "my-db")
    {
        ExcelExporter.Generate(outputPath, sql, server, database);
        return MashupReader.ReadSection1(outputPath);
    }

    [Fact]
    public void Writes_server_database_and_sql_into_the_query()
    {
        string m = Generate("SELECT * FROM leagues", "srv.example.net", "sales");

        Assert.Contains("\"srv.example.net\"", m);
        Assert.Contains("\"sales\"", m);
        Assert.Contains("SELECT * FROM leagues", m);
    }

    [Fact]
    public void Removes_the_template_placeholders()
    {
        string m = Generate("SELECT 1");

        Assert.DoesNotContain("fomf-sandbox-ssdb", m);
        Assert.DoesNotContain("fomf-ss-sandbox-db", m);
        Assert.DoesNotContain("Refresh All", m);
    }

    [Fact]
    public void Escapes_double_quotes_for_M()
    {
        string m = Generate("SELECT name AS \"League Name\" FROM leagues");

        Assert.Contains("SELECT name AS \"\"League Name\"\" FROM leagues", m);
    }

    [Fact]
    public void Escapes_quotes_in_server_and_database()
    {
        string m = Generate("SELECT 1", "odd\"server", "odd\"db");

        Assert.Contains("odd\"\"server", m);
        Assert.Contains("odd\"\"db", m);
    }

    [Fact]
    public void Keeps_multiline_sql_intact()
    {
        string sql = "-- leagues\r\nSELECT id,\r\n       name\r\nFROM leagues";

        string m = Generate(sql);

        Assert.Contains(sql, m);
    }

    [Fact]
    public void Produces_a_readable_workbook()
    {
        Generate("SELECT 1");

        using var zip = ZipFile.OpenRead(outputPath);
        Assert.NotNull(zip.GetEntry("xl/workbook.xml"));
        Assert.NotNull(zip.GetEntry("[Content_Types].xml"));
    }

    [Fact]
    public void Overwrites_an_existing_file()
    {
        File.WriteAllText(outputPath, "not a workbook");

        string m = Generate("SELECT 2");

        Assert.Contains("SELECT 2", m);
    }
}
