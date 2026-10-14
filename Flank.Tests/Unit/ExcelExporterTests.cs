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

    private static readonly System.Xml.Linq.XNamespace Sheet =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private System.Xml.Linq.XDocument ReadPart(string path)
    {
        using var zip = ZipFile.OpenRead(outputPath);
        using var stream = zip.GetEntry(path)!.Open();
        return System.Xml.Linq.XDocument.Load(stream);
    }

    private string GenerateWithParameters(params string[] names)
    {
        ExcelExporter.Generate(
            outputPath,
            "EXEC [dbo].[get_things] " + string.Join(", ", names.Select(n => $"@{n} = @{n}")),
            "my-server",
            "my-db",
            names);
        return MashupReader.ReadSection1(outputPath);
    }

    [Fact]
    public void Parameters_are_passed_to_the_query_as_sql_parameters()
    {
        string m = GenerateWithParameters("League_Id", "Name");

        Assert.Contains("Excel.CurrentWorkbook(){[Name=\"FlankParameters\"]}[Content]", m);
        Assert.Contains("Record.FromList(Parameters[Column2], Parameters[Column1])", m);
        Assert.Contains(
            "Value.NativeQuery(Sql.Database(\"my-server\", \"my-db\"), " +
            "\"EXEC [dbo].[get_things] @League_Id = @League_Id, @Name = @Name\", Values)",
            m);
    }

    [Fact]
    public void Parameter_names_go_in_column_a_above_the_table()
    {
        GenerateWithParameters("League_Id", "Name");

        // Note (1), last refreshed (2), gap (3), parameters (4-5), gap (6), table (7).
        var rows = ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "row").ToList();
        Assert.Equal(new[] { "1", "2", "4", "5", "7" }, rows.Select(r => (string)r.Attribute("r")!));
        Assert.Equal("League_Id", (string)rows[2].Descendants(Sheet + "t").Single());
        Assert.Equal("Name", (string)rows[3].Descendants(Sheet + "t").Single());
        Assert.Equal("A7", (string)rows[4].Element(Sheet + "c")!.Attribute("r")!);

        var table = ReadPart("xl/tables/table1.xml").Root!;
        Assert.Equal("A7:A8", (string)table.Attribute("ref")!);
        Assert.Equal("A7:A8", (string)table.Element(Sheet + "autoFilter")!.Attribute("ref")!);
    }

    [Fact]
    public void Parameter_cells_are_named_and_the_table_overwrites_instead_of_shifting_them()
    {
        GenerateWithParameters("League_Id", "Name");

        var names = ReadPart("xl/workbook.xml").Descendants(Sheet + "definedName")
            .ToDictionary(n => (string)n.Attribute("name")!, n => n.Value);
        Assert.Equal("Query1!$A$4:$B$5", names["FlankParameters"]);
        Assert.Equal("Query1!$A$7:$A$8", names["ExternalData_1"]);

        var queryTable = ReadPart("xl/queryTables/queryTable1.xml").Root!;
        Assert.Equal("overwriteClear", (string)queryTable.Attribute("growShrinkType")!);
    }

    [Fact]
    public void Known_columns_are_created_up_front_below_the_parameters()
    {
        ExcelExporter.Generate(
            outputPath, "EXEC [dbo].[x] @League_Id = @League_Id", "s", "d",
            new[] { "League_Id" },
            new[] { "team_name", "team_id", "league_id" });

        // One parameter row (4), a gap, then the table at row 6.
        var table = ReadPart("xl/tables/table1.xml").Root!;
        Assert.Equal("A6:C7", (string)table.Attribute("ref")!);
        Assert.Equal(
            new[] { "team_name", "team_id", "league_id" },
            table.Descendants(Sheet + "tableColumn").Select(c => (string)c.Attribute("name")!));

        var fields = ReadPart("xl/queryTables/queryTable1.xml").Descendants(Sheet + "queryTableField").ToList();
        Assert.Equal(new[] { "team_name", "team_id", "league_id" }, fields.Select(f => (string)f.Attribute("name")!));
        Assert.Equal(new[] { "1", "2", "3" }, fields.Select(f => (string)f.Attribute("tableColumnId")!));

        var header = ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "row").Single(r => (string)r.Attribute("r")! == "6");
        Assert.Equal(new[] { "team_name", "team_id", "league_id" }, header.Descendants(Sheet + "t").Select(t => (string)t));

        var names = ReadPart("xl/workbook.xml").Descendants(Sheet + "definedName")
            .ToDictionary(n => (string)n.Attribute("name")!, n => n.Value);
        Assert.Equal("Query1!$A$6:$C$7", names["ExternalData_1"]);
        Assert.Equal("Query1!$A$4:$B$4", names["FlankParameters"]);
    }

    [Fact]
    public void Colliding_column_names_leave_the_placeholder_for_excel_to_name()
    {
        ExcelExporter.Generate(outputPath, "EXEC x", "s", "d", Array.Empty<string>(), new[] { "id", "ID" });

        var table = ReadPart("xl/tables/table1.xml").Root!;
        Assert.Equal("A4:A5", (string)table.Attribute("ref")!);
        Assert.Contains("Refresh All", (string)table.Descendants(Sheet + "tableColumn").Single().Attribute("name")!);
    }

    [Fact]
    public void Without_parameters_the_layout_is_unchanged()
    {
        Generate("SELECT 1");

        Assert.Equal("A4:A5", (string)ReadPart("xl/tables/table1.xml").Root!.Attribute("ref")!);
        Assert.DoesNotContain(
            ReadPart("xl/workbook.xml").Descendants(Sheet + "definedName"),
            n => (string)n.Attribute("name")! == "FlankParameters");
        Assert.DoesNotContain("Value.NativeQuery", MashupReader.ReadSection1(outputPath));
    }

    private string Cell(string reference)
    {
        var cell = ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "c")
            .Single(c => (string)c.Attribute("r")! == reference);

        // t="s" cells hold an index into the shared strings.
        if ((string?)cell.Attribute("t") == "s")
        {
            int index = (int)cell.Element(Sheet + "v")!;
            var item = ReadPart("xl/sharedStrings.xml").Root!.Elements(Sheet + "si").ElementAt(index);
            return string.Concat(item.Descendants(Sheet + "t").Select(t => (string)t));
        }

        return string.Concat(cell.Descendants(Sheet + "t").Select(t => (string)t));
    }

    [Fact]
    public void Starts_with_a_how_to_refresh_note()
    {
        Generate("SELECT 1");

        Assert.StartsWith("To update: Data > Refresh All.", Cell("A1"));
        Assert.DoesNotContain("privacy", Cell("A1"));
    }

    [Fact]
    public void Parameter_workbooks_explain_entering_values_and_the_privacy_prompt()
    {
        GenerateWithParameters("League_Id");

        Assert.StartsWith("To update: enter values in column B", Cell("A1"));
        Assert.Contains("Ignore Privacy Levels", Cell("A1"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Keeps_the_last_refreshed_query(bool withParameters)
    {
        string m = withParameters ? GenerateWithParameters("League_Id") : Generate("SELECT 1");

        Assert.Contains("shared LastRefreshed = let", m);
        Assert.Contains("DateTime.LocalNow()", m);
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(m, "^shared ", System.Text.RegularExpressions.RegexOptions.Multiline).Count);
        Assert.Equal("Last refreshed: never", Cell("A2"));
    }

    [Fact]
    public void Overwrites_an_existing_file()
    {
        File.WriteAllText(outputPath, "not a workbook");

        string m = Generate("SELECT 2");

        Assert.Contains("SELECT 2", m);
    }
}
