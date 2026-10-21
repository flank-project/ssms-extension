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

        // Each cell by its name, so inserting rows doesn't break the query.
        Assert.Contains("Excel.CurrentWorkbook(){[Name=\"Flank_League_Id\"]}[Content]{0}[Column1]", m);
        Assert.Contains("Excel.CurrentWorkbook(){[Name=\"Flank_Name\"]}[Content]{0}[Column1]", m);
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

        Assert.Equal("League", Text(5, "A"));
        Assert.Equal("Team name", Text(6, "A"));

        // A bold header row above the parameters; notes aren't a column.
        Assert.Equal("Parameter", Text(4, "A"));
        Assert.Equal("User Input", Text(4, "B"));
        Assert.Equal("", Text(4, "C"));
        Assert.Equal("", Text(5, "C"));

        // The note is a cell note on the label, which moves with it.
        var comment = ReadPart("xl/comments1.xml").Descendants(Sheet + "comment").Single();
        Assert.Equal("A5", (string)comment.Attribute("ref")!);
        Assert.Equal("1 = NBA, 2 = NFL", comment.Descendants(Sheet + "t").Single().Value);

        var sheet = ReadPart("xl/worksheets/sheet1.xml").Root!;
        var order = sheet.Elements().Select(e => e.Name.LocalName).ToList();
        Assert.True(order.IndexOf("legacyDrawing") < order.IndexOf("tableParts"));
        var relationships = ReadPart("xl/worksheets/_rels/sheet1.xml.rels").Root!.Elements().Select(r => (string)r.Attribute("Target")!).ToList();
        Assert.Contains("../comments1.xml", relationships);
        Assert.Contains("../drawings/vmlDrawing1.vml", relationships);
        using (var zip = ZipFile.OpenRead(outputPath))
        {
            using var vml = new StreamReader(zip.GetEntry("xl/drawings/vmlDrawing1.vml")!.Open());
            string drawing = vml.ReadToEnd();
            Assert.Contains("<x:Row>4</x:Row>", drawing);
            Assert.Contains("<x:Column>0</x:Column>", drawing);
        }

        var types = ReadPart("[Content_Types].xml").ToString();
        Assert.Contains("Extension=\"vml\"", types);
        Assert.Contains("/xl/comments1.xml", types);

        string m = MashupReader.ReadSection1(outputPath);
        Assert.Contains("{\"League_Id\", \"Name\", \"Active\"}", m);
    }

    [Fact]
    public void Reviewed_parameters_get_data_validation_and_required_highlighting()
    {
        ExcelExporter.Generate(outputPath, ReviewedRequest());

        var sheet = ReadPart("xl/worksheets/sheet1.xml").Root!;
        var validations = sheet.Descendants(Sheet + "dataValidation").ToDictionary(v => (string)v.Attribute("sqref")!);

        Assert.Equal("whole", (string)validations["B5"].Attribute("type")!);
        Assert.Equal("2147483647", (string)validations["B5"].Element(Sheet + "formula2")!);
        Assert.Equal("1 = NBA, 2 = NFL", (string)validations["B5"].Attribute("prompt")!);
        Assert.Equal("textLength", (string)validations["B6"].Attribute("type")!);
        Assert.Equal("50", (string)validations["B6"].Element(Sheet + "formula1")!);
        Assert.Equal("list", (string)validations["B7"].Attribute("type")!);
        Assert.Equal("\"TRUE,FALSE\"", (string)validations["B7"].Element(Sheet + "formula1")!);

        // Only the required cell is highlighted while empty.
        var format = sheet.Elements(Sheet + "conditionalFormatting").Single();
        Assert.Equal("B5", (string)format.Attribute("sqref")!);
        Assert.Equal("LEN(TRIM(B5))=0", (string)format.Descendants(Sheet + "formula").Single());

        // Schema order: right after sheetData, validations after formatting, before pageMargins.
        var order = sheet.Elements().Select(e => e.Name.LocalName).ToList();
        Assert.True(order.IndexOf("sheetData") < order.IndexOf("conditionalFormatting"));
        Assert.True(order.IndexOf("conditionalFormatting") < order.IndexOf("dataValidations"));
        Assert.True(order.IndexOf("dataValidations") < order.IndexOf("pageMargins"));

        var dxfs = ReadPart("xl/styles.xml").Root!.Element(Sheet + "dxfs")!;
        Assert.Equal(dxfs.Elements(Sheet + "dxf").Count().ToString(), (string)dxfs.Attribute("count")!);

        // Refreshing mustn't narrow the parameter column.
        Assert.Equal("0", (string)ReadPart("xl/queryTables/queryTable1.xml").Root!.Attribute("adjustColumnWidth")!);
        Assert.Equal("0", (string)ReadPart("xl/queryTables/queryTable2.xml").Root!.Attribute("adjustColumnWidth")!);
    }

    private static readonly System.Xml.Linq.XNamespace FeaturePropertyBag =
        "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag";

    private System.Xml.Linq.XElement ParameterCell(string reference) =>
        ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "c").Single(c => (string)c.Attribute("r")! == reference);

    // The cell format's complement index, if it's a checkbox.
    private string? CheckboxComplement(System.Xml.Linq.XElement cell)
    {
        var xf = ReadPart("xl/styles.xml").Root!.Element(Sheet + "cellXfs")!.Elements(Sheet + "xf")
            .ElementAt((int?)cell.Attribute("s") ?? 0);
        return (string?)xf.Descendants(FeaturePropertyBag + "xfComplement").SingleOrDefault()?.Attribute("i");
    }

    [Fact]
    public void Bit_parameters_are_checkboxes_with_the_dropdown_kept_for_older_excel()
    {
        ExcelExporter.Generate(outputPath, ReviewedRequest());

        var active = ParameterCell("B7");
        Assert.Equal("b", (string)active.Attribute("t")!);
        Assert.Equal("1", (string)active.Element(Sheet + "v")!);
        Assert.Equal("0", CheckboxComplement(active));
        Assert.Null(CheckboxComplement(ParameterCell("B5")));

        // Complement 0 is the checkbox control, as Excel saves it.
        var bags = ReadPart("xl/featurePropertyBag/featurePropertyBag.xml").Root!.Elements(FeaturePropertyBag + "bag").ToList();
        Assert.Equal(new[] { "Checkbox", "XFControls", "XFComplement", "XFComplements" }, bags.Select(b => (string)b.Attribute("type")!));
        Assert.Equal("XFComplementsMapperExtRef", (string)bags[3].Attribute("extRef")!);

        var relationship = ReadPart("xl/_rels/workbook.xml.rels").Root!.Elements()
            .Single(r => (string)r.Attribute("Type")! == "http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag");
        Assert.Equal("featurePropertyBag/featurePropertyBag.xml", (string)relationship.Attribute("Target")!);

        Assert.Contains(ReadPart("[Content_Types].xml").Root!.Elements(),
            o => (string?)o.Attribute("PartName") == "/xl/featurePropertyBag/featurePropertyBag.xml" &&
                 (string?)o.Attribute("ContentType") == "application/vnd.ms-excel.featurepropertybag+xml");

        var validation = ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "dataValidation")
            .Single(v => (string)v.Attribute("sqref")! == "B7");
        Assert.Equal("\"TRUE,FALSE\"", (string)validation.Element(Sheet + "formula1")!);
    }

    [Fact]
    public void A_blank_bit_parameter_is_a_blank_checkbox_not_false()
    {
        var request = ReviewedRequest();
        request.ParameterValues = new object?[] { 2L, null, null };

        ExcelExporter.Generate(outputPath, request);

        // Shown greyed out in Excel; the query reads it as blank (the procedure's default, or NULL).
        var active = ParameterCell("B7");
        Assert.Null(active.Element(Sheet + "v"));
        Assert.Equal("0", CheckboxComplement(active));
    }

    [Fact]
    public void Workbooks_without_bit_parameters_have_no_checkbox_part()
    {
        GenerateWithParameters("League_Id");

        using var zip = ZipFile.OpenRead(outputPath);
        Assert.Null(zip.GetEntry("xl/featurePropertyBag/featurePropertyBag.xml"));
    }

    [Fact]
    public void The_query_checks_values_again_since_pasting_skips_validation()
    {
        ExcelExporter.Generate(outputPath, ReviewedRequest());
        string m = MashupReader.ReadSection1(outputPath);

        // Named by label, not cell: inserting rows would make a cell address wrong.
        Assert.Contains("\"League is required. Enter a value, then refresh.\"", m);
        Assert.Contains("\"League must be a whole number between -2,147,483,648 and 2,147,483,647.\"", m);
        Assert.Contains("\"Team name can be at most 50 characters.\"", m);
        Assert.Contains("\"Active only must be TRUE or FALSE.\"", m);
    }

    private WorkbookRequest ListRequest() =>
        new WorkbookRequest("EXEC [dbo].[p] @League_Id = @League_Id", "my-server", "my-db")
        {
            Parameters = new[]
            {
                new WorkbookParameter("League_Id", "League")
                {
                    Required = true,
                    List = new WorkbookList("EXEC [dbo].[get_leagues]", "id", "name")
                    {
                        Rows = new (object?, object?)[] { (1, "NBA"), (2, "NFL"), (3, "MLB") },
                    },
                },
            },
            ParameterValues = new object?[] { 2 },
        };

    [Fact]
    public void A_list_parameter_shows_the_label_and_a_dropdown_from_a_hidden_lists_sheet()
    {
        ExcelExporter.Generate(outputPath, ListRequest());

        var sheet = ReadPart("xl/worksheets/sheet1.xml").Root!;
        var b4 = sheet.Descendants(Sheet + "c").Single(c => (string)c.Attribute("r")! == "B5");
        Assert.Equal("NFL", b4.Descendants(Sheet + "t").Single().Value);

        var validation = sheet.Descendants(Sheet + "dataValidation").Single();
        Assert.Equal("list", (string)validation.Attribute("type")!);
        Assert.Equal("FlankList1", (string)validation.Element(Sheet + "formula1")!);

        var book = ReadPart("xl/workbook.xml").Root!;
        var lists = book.Descendants(Sheet + "sheet").Single(s => (string)s.Attribute("name")! == "DropdownOptions");
        Assert.Equal("hidden", (string)lists.Attribute("state")!);
        var names = book.Descendants(Sheet + "definedName").ToDictionary(n => (string)n.Attribute("name")!, n => n.Value);
        Assert.Equal("League_options[name]", names["FlankList1"]);
        Assert.Equal("DropdownOptions!$A$1:$B$4", names["ExternalData_3"]);

        var table = ReadPart("xl/tables/table3.xml").Root!;
        Assert.Equal("League_options", (string)table.Attribute("name")!);
        Assert.Equal(new[] { "name", "id" }, table.Descendants(Sheet + "tableColumn").Select(c => (string)c.Attribute("name")!));

        var listSheet = ReadPart("xl/worksheets/sheet2.xml").Root!;
        Assert.Equal(4, listSheet.Descendants(Sheet + "row").Count());

        var connection = ReadPart("xl/connections.xml").Root!.Elements(Sheet + "connection").Single(c => (string)c.Attribute("id")! == "3");
        Assert.Contains("Location=League_options", (string)connection.Element(Sheet + "dbPr")!.Attribute("connection")!);

        var types = ReadPart("[Content_Types].xml").ToString();
        Assert.Contains("/xl/worksheets/sheet2.xml", types);
        Assert.Contains("/xl/tables/table3.xml", types);
        Assert.Contains("/xl/queryTables/queryTable3.xml", types);
    }

    [Fact]
    public void Lists_are_named_after_their_parameters_and_kept_unique()
    {
        WorkbookParameter Listed(string sqlName, string displayName) =>
            new WorkbookParameter(sqlName, displayName)
            {
                List = new WorkbookList("EXEC [dbo].[get_leagues]", "id", "name")
                {
                    Rows = new (object?, object?)[] { (1, "NBA") },
                },
            };

        ExcelExporter.Generate(outputPath, new WorkbookRequest("EXEC [dbo].[p] @a = @a, @b = @b", "s", "d")
        {
            Parameters = new[] { Listed("a", "Home team"), Listed("b", "Home team") },
        });

        string m = MashupReader.ReadSection1(outputPath);
        Assert.Contains("shared Home_team_options = let", m);
        Assert.Contains("shared Home_team_options_2 = let", m);
        Assert.Equal("Home_team_options_2", (string)ReadPart("xl/tables/table4.xml").Root!.Attribute("name")!);
    }

    [Fact]
    public void The_list_is_its_own_query_and_the_main_query_looks_up_the_value()
    {
        ExcelExporter.Generate(outputPath, ListRequest());
        string m = MashupReader.ReadSection1(outputPath);

        Assert.Contains("shared League_options = let", m);
        Assert.Contains("[Query=\"EXEC [dbo].[get_leagues]\"]", m);
        Assert.Contains("Table.SelectColumns(Source, {\"name\", \"id\"})", m);
        Assert.Contains("Table.SelectRows(League_options, each Text.From(Record.Field(_, \"name\")) = Text.From(v))", m);
        Assert.Contains("Record.Field(match{0}, \"id\")", m);
        Assert.Contains("\"League must be one of the choices in its list.\"", m);
    }

    [Fact]
    public void Parameter_names_go_in_column_a_above_the_table()
    {
        GenerateWithParameters("League_Id", "Name");

        // Note (1), last refreshed (2), gap (3), header (4), parameters (5-6), gap (7), table (8).
        var rows = ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "row").ToList();
        Assert.Equal(new[] { "1", "2", "4", "5", "6", "8" }, rows.Select(r => (string)r.Attribute("r")!));
        Assert.Equal(new[] { "Parameter", "User Input" }, rows[2].Descendants(Sheet + "t").Select(t => (string)t));
        Assert.Equal("League_Id", (string)rows[3].Descendants(Sheet + "t").Single());
        Assert.Equal("Name", (string)rows[4].Descendants(Sheet + "t").Single());
        Assert.Equal("A8", (string)rows[5].Element(Sheet + "c")!.Attribute("r")!);

        // The header is bold.
        int style = (int)rows[2].Elements(Sheet + "c").First().Attribute("s")!;
        var styles = ReadPart("xl/styles.xml").Root!;
        int font = (int)styles.Element(Sheet + "cellXfs")!.Elements(Sheet + "xf").ElementAt(style).Attribute("fontId")!;
        Assert.NotNull(styles.Element(Sheet + "fonts")!.Elements(Sheet + "font").ElementAt(font).Element(Sheet + "b"));

        var table = ReadPart("xl/tables/table1.xml").Root!;
        Assert.Equal("A8:A9", (string)table.Attribute("ref")!);
        Assert.Equal("A8:A9", (string)table.Element(Sheet + "autoFilter")!.Attribute("ref")!);
    }

    [Fact]
    public void Parameter_cells_are_named_and_the_table_overwrites_instead_of_shifting_them()
    {
        GenerateWithParameters("League_Id", "Name");

        var names = ReadPart("xl/workbook.xml").Descendants(Sheet + "definedName")
            .ToDictionary(n => (string)n.Attribute("name")!, n => n.Value);
        Assert.Equal("Query1!$B$5", names["Flank_League_Id"]);
        Assert.Equal("Query1!$B$6", names["Flank_Name"]);
        Assert.Equal("Query1!$A$8:$A$9", names["ExternalData_1"]);

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

        // The header row (4), one parameter row (5), a gap, then the table at row 7.
        var table = ReadPart("xl/tables/table1.xml").Root!;
        Assert.Equal("A7:C8", (string)table.Attribute("ref")!);
        Assert.Equal(
            new[] { "team_name", "team_id", "league_id" },
            table.Descendants(Sheet + "tableColumn").Select(c => (string)c.Attribute("name")!));

        var fields = ReadPart("xl/queryTables/queryTable1.xml").Descendants(Sheet + "queryTableField").ToList();
        Assert.Equal(new[] { "team_name", "team_id", "league_id" }, fields.Select(f => (string)f.Attribute("name")!));
        Assert.Equal(new[] { "1", "2", "3" }, fields.Select(f => (string)f.Attribute("tableColumnId")!));

        var header = ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "row").Single(r => (string)r.Attribute("r")! == "7");
        Assert.Equal(new[] { "team_name", "team_id", "league_id" }, header.Descendants(Sheet + "t").Select(t => (string)t));

        var names = ReadPart("xl/workbook.xml").Descendants(Sheet + "definedName")
            .ToDictionary(n => (string)n.Attribute("name")!, n => n.Value);
        Assert.Equal("Query1!$A$7:$C$8", names["ExternalData_1"]);
        Assert.Equal("Query1!$B$5", names["Flank_League_Id"]);
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
            n => ((string)n.Attribute("name")!).StartsWith("Flank_"));
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

        Assert.Equal("1", (string)CellElement("B5").Element(Sheet + "v")!);
        Assert.Equal("abc", Cell("B6"));
        Assert.DoesNotContain(
            ReadPart("xl/worksheets/sheet1.xml").Descendants(Sheet + "c"),
            c => (string)c.Attribute("r")! == "B7");
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

        // Parameter header at row 4, parameter at 5, gap, table header at 7, data 8-10.
        Assert.Equal("A7:A10", (string)ReadPart("xl/tables/table1.xml").Root!.Attribute("ref")!);
        Assert.Equal("3", (string)CellElement("A10").Element(Sheet + "v")!);
        Assert.Equal(
            new[] { "1", "2", "4", "5", "7", "8", "9", "10" },
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
