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

        Assert.Contains("Excel.CurrentWorkbook(){[Name=\"FlankParameters\"]}[Content][Column2]", m);
        // By position, so the labels in column A can be anything.
        Assert.Contains("Record.FromList({Value0, Value1}, {\"League_Id\", \"Name\"})", m);
        Assert.Contains(
            "Value.NativeQuery(Sql.Database(\"my-server\", \"my-db\"), " +
            "\"EXEC [dbo].[get_things] @League_Id = @League_Id, @Name = @Name\", Values)",
            m);
    }

    private WorkbookRequest ReviewedRequest() =>
        new WorkbookRequest("EXEC [dbo].[p] @League_Id = @League_Id, @Name = @Name, @Active = @Active", "my-server", "my-db")
        {
            Parameters = new[]
            {
                new WorkbookParameter("League_Id", "League")
                {
                    Kind = WorkbookParameterKind.WholeNumber, Minimum = -2147483648, Maximum = 2147483647,
                    Required = true, Note = "1 = NBA, 2 = NFL",
                },
                new WorkbookParameter("Name", "Team name") { Kind = WorkbookParameterKind.Text, MaxLength = 50 },
                new WorkbookParameter("Active", "Active only") { Kind = WorkbookParameterKind.TrueFalse },
            },
            ParameterValues = new object?[] { 2L, null, true },
        };

    [Fact]
    public void Reviewed_parameters_show_their_display_names_and_notes()
    {
        ExcelExporter.Generate(outputPath, ReviewedRequest());

        var rows = ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "row").ToList();
        string Text(int row, string column) =>
            rows.Single(r => (string)r.Attribute("r")! == row.ToString())
                .Elements(Sheet + "c").SingleOrDefault(c => (string)c.Attribute("r")! == column + row)?
                .Descendants(Sheet + "t").SingleOrDefault()?.Value ?? "";

        Assert.Equal("League", Text(4, "A"));
        Assert.Equal("Required. 1 = NBA, 2 = NFL", Text(4, "C"));
        Assert.Equal("Team name", Text(5, "A"));
        Assert.Equal("", Text(5, "C"));

        string m = MashupReader.ReadSection1(outputPath);
        Assert.Contains("{\"League_Id\", \"Name\", \"Active\"}", m);
    }

    [Fact]
    public void Reviewed_parameters_get_data_validation_and_required_highlighting()
    {
        ExcelExporter.Generate(outputPath, ReviewedRequest());

        var sheet = ReadPart("xl/worksheets/sheet1.xml").Root!;
        var validations = sheet.Descendants(Sheet + "dataValidation").ToDictionary(v => (string)v.Attribute("sqref")!);

        Assert.Equal("whole", (string)validations["B4"].Attribute("type")!);
        Assert.Equal("2147483647", (string)validations["B4"].Element(Sheet + "formula2")!);
        Assert.Equal("1 = NBA, 2 = NFL", (string)validations["B4"].Attribute("prompt")!);
        Assert.Equal("textLength", (string)validations["B5"].Attribute("type")!);
        Assert.Equal("50", (string)validations["B5"].Element(Sheet + "formula1")!);
        Assert.Equal("list", (string)validations["B6"].Attribute("type")!);
        Assert.Equal("\"TRUE,FALSE\"", (string)validations["B6"].Element(Sheet + "formula1")!);

        // Only the required cell is highlighted while empty.
        var format = sheet.Elements(Sheet + "conditionalFormatting").Single();
        Assert.Equal("B4", (string)format.Attribute("sqref")!);
        Assert.Equal("LEN(TRIM(B4))=0", (string)format.Descendants(Sheet + "formula").Single());

        // Schema order: right after sheetData, validations after formatting, before pageMargins.
        var order = sheet.Elements().Select(e => e.Name.LocalName).ToList();
        Assert.True(order.IndexOf("sheetData") < order.IndexOf("conditionalFormatting"));
        Assert.True(order.IndexOf("conditionalFormatting") < order.IndexOf("dataValidations"));
        Assert.True(order.IndexOf("dataValidations") < order.IndexOf("pageMargins"));

        var dxfs = ReadPart("xl/styles.xml").Root!.Element(Sheet + "dxfs")!;
        Assert.Equal(dxfs.Elements(Sheet + "dxf").Count().ToString(), (string)dxfs.Attribute("count")!);
    }

    [Fact]
    public void The_query_checks_values_again_since_pasting_skips_validation()
    {
        ExcelExporter.Generate(outputPath, ReviewedRequest());
        string m = MashupReader.ReadSection1(outputPath);

        Assert.Contains("League (cell B4) is required. Enter a value, then refresh.", m);
        Assert.Contains("League (cell B4) must be a whole number between -2,147,483,648 and 2,147,483,647.", m);
        Assert.Contains("Team name (cell B5) can be at most 50 characters.", m);
        Assert.Contains("Active only (cell B6) must be TRUE or FALSE.", m);
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

    private System.Xml.Linq.XElement CellElement(string reference) =>
        ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "c")
            .Single(c => (string)c.Attribute("r")! == reference);

    [Fact]
    public void Parameter_values_fill_column_b()
    {
        ExcelExporter.Generate(outputPath, new WorkbookRequest("EXEC x @a = @a, @b = @b, @c = @c", "s", "d")
        {
            ParameterNames = new[] { "a", "b", "c" },
            ParameterValues = new object?[] { 1, "abc", null },
        });

        Assert.Equal("1", (string)CellElement("B4").Element(Sheet + "v")!);
        Assert.Equal("abc", Cell("B5"));
        Assert.DoesNotContain(
            ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "c"),
            c => (string)c.Attribute("r")! == "B6");
    }

    [Fact]
    public void Preview_rows_fill_the_table_with_typed_cells()
    {
        ExcelExporter.Generate(outputPath, new WorkbookRequest("SELECT 1", "s", "d")
        {
            ColumnNames = new[] { "id", "name", "played", "active" },
            Rows = new IReadOnlyList<object?>[]
            {
                new object?[] { 1, "NBA", new DateTime(2024, 1, 2, 3, 4, 0), true },
                new object?[] { 2.5m, null, DBNull.Value, false },
            },
        });

        var table = ReadPart("xl/tables/table1.xml").Root!;
        Assert.Equal("A4:D6", (string)table.Attribute("ref")!);

        var names = ReadPart("xl/workbook.xml").Descendants(Sheet + "definedName")
            .ToDictionary(n => (string)n.Attribute("name")!, n => n.Value);
        Assert.Equal("Query1!$A$4:$D$6", names["ExternalData_1"]);

        Assert.Equal("1", (string)CellElement("A5").Element(Sheet + "v")!);
        Assert.Equal("NBA", Cell("B5"));
        Assert.Equal("b", (string)CellElement("D5").Attribute("t")!);
        Assert.Equal("1", (string)CellElement("D5").Element(Sheet + "v")!);
        Assert.Equal("2.5", (string)CellElement("A6").Element(Sheet + "v")!);
        Assert.Equal("0", (string)CellElement("D6").Element(Sheet + "v")!);

        // Dates are serial numbers with a date format.
        var date = CellElement("C5");
        Assert.Equal(new DateTime(2024, 1, 2, 3, 4, 0).ToOADate(), double.Parse((string)date.Element(Sheet + "v")!, System.Globalization.CultureInfo.InvariantCulture));
        int style = (int)date.Attribute("s")!;
        var xf = ReadPart("xl/styles.xml").Root!.Element(Sheet + "cellXfs")!.Elements(Sheet + "xf").ElementAt(style);
        Assert.Equal("22", (string)xf.Attribute("numFmtId")!);

        // Nulls leave the cell out.
        Assert.DoesNotContain(
            ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "c"),
            c => (string)c.Attribute("r")! == "B6" || (string)c.Attribute("r")! == "C6");
    }

    [Fact]
    public void Preview_rows_show_when_they_were_fetched()
    {
        ExcelExporter.Generate(outputPath, new WorkbookRequest("SELECT 1", "s", "d")
        {
            ColumnNames = new[] { "id" },
            Rows = new IReadOnlyList<object?>[] { new object?[] { 1 } },
            RowsFetchedAt = new DateTime(2026, 10, 14, 2, 56, 0),
        });

        Assert.Equal("Last refreshed: 2026-10-14 02:56", Cell("A2"));
    }

    [Fact]
    public void Preview_rows_go_below_the_parameters()
    {
        ExcelExporter.Generate(outputPath, new WorkbookRequest("EXEC x @a = @a", "s", "d")
        {
            ParameterNames = new[] { "a" },
            ParameterValues = new object?[] { 7 },
            ColumnNames = new[] { "id" },
            Rows = new IReadOnlyList<object?>[] { new object?[] { 1 }, new object?[] { 2 }, new object?[] { 3 } },
        });

        // Parameter at row 4, gap, header at 6, data 7-9.
        Assert.Equal("A6:A9", (string)ReadPart("xl/tables/table1.xml").Root!.Attribute("ref")!);
        Assert.Equal("3", (string)CellElement("A9").Element(Sheet + "v")!);
        Assert.Equal(
            new[] { "1", "2", "4", "6", "7", "8", "9" },
            ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "row").Select(r => (string)r.Attribute("r")!));
    }

    [Fact]
    public void Has_a_single_sheet()
    {
        Generate("SELECT 1");

        var sheets = ReadPart("xl/workbook.xml").Descendants(Sheet + "sheet").ToList();
        Assert.Single(sheets);
    }

    [Fact]
    public void Overwrites_an_existing_file()
    {
        File.WriteAllText(outputPath, "not a workbook");

        string m = Generate("SELECT 2");

        Assert.Contains("SELECT 2", m);
    }
}
