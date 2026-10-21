using Flank.Excel;
using Flank.Ssrs;
using Flank.Tests.Support;
using System.IO.Compression;
using System.Xml.Linq;

namespace Flank.Tests.Unit;

/// <summary>Table-valued parameters: the call that passes them as JSON, the workbook's tables, and reports.</summary>
public class TableParameterTests : IDisposable
{
    private static readonly XNamespace Sheet =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly XNamespace Rdl =
        "http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition";

    private readonly string outputPath =
        Path.Combine(Path.GetTempPath(), $"flank-unit-{Guid.NewGuid():N}.xlsx");

    public void Dispose()
    {
        File.Delete(outputPath);
    }

    private static SprocParameter Column(string name, string type, bool nullable = true, short maxLength = 0) =>
        new SprocParameter { Name = name, SqlType = type, IsNullable = nullable, MaxLength = maxLength };

    private static SprocParameter Teams() => new SprocParameter
    {
        Name = "@Teams",
        SqlType = "team_list",
        MaxLength = -1,
        TableType = "[dbo].[team_list]",
        Columns = new[] { Column("team_id", "int", nullable: false), Column("team_name", "nvarchar", maxLength: 100) },
    };

    // ---------------------------------------------------------------- the call

    [Fact]
    public void The_call_fills_a_table_variable_from_json_and_passes_it()
    {
        var parameters = new[] { new SprocParameter { Name = "@League_Id", SqlType = "int" }, Teams() };

        string call = ProcedureCall.BuildExec("[dbo].[p]", parameters);

        Assert.Equal(
            "DECLARE @flank_table_Teams [dbo].[team_list];\r\n" +
            "INSERT INTO @flank_table_Teams ([team_id], [team_name]) SELECT [c0], [c1] " +
            "FROM OPENJSON(@Teams) WITH ([c0] int '$[0]', [c1] nvarchar(50) '$[1]');\r\n" +
            "EXEC [dbo].[p] @League_Id = @League_Id, @Teams = @flank_table_Teams",
            call);
    }

    [Fact]
    public void A_table_is_passed_through_sp_executesql_when_other_parameters_may_be_left_out()
    {
        var parameters = new[] { new SprocParameter { Name = "@Since", SqlType = "date", HasDefault = true }, Teams() };
        var settings = ParameterSettings.FromProcedure(parameters);

        // A table is never left out: with no rows it's an empty table.
        settings[1].IfBlank = BlankValue.UseProcedureDefault;

        string call = ProcedureCall.BuildExec("[dbo].[p]", parameters, settings);

        Assert.StartsWith("DECLARE @flank_table_Teams [dbo].[team_list];", call);
        Assert.Contains("CASE WHEN @Since IS NOT NULL THEN N', @Since = @Since' END", call);
        Assert.Contains("N', @Teams = @Teams'", call);
        Assert.Contains("N'@Since date, @Teams [dbo].[team_list] READONLY'", call);
        Assert.EndsWith("@Since = @Since, @Teams = @flank_table_Teams;", call);
    }

    [Fact]
    public void Rows_become_json_arrays_in_column_order()
    {
        var rows = new IReadOnlyList<object?>[]
        {
            new object?[] { 1, "a \"quoted\" \\ name", true, new DateTime(2024, 1, 2), 12.5m },
            new object?[] { 2L, null, false, new DateTime(2024, 1, 2, 10, 11, 12, 500), DBNull.Value },
        };

        Assert.Equal(
            "[[1,\"a \\\"quoted\\\" \\\\ name\",true,\"2024-01-02T00:00:00\",12.5]," +
            "[2,null,false,\"2024-01-02T10:11:12.5\",null]]",
            TableValues.ToJson(rows));
        Assert.Null(TableValues.ToJson(Array.Empty<IReadOnlyList<object?>>()));
    }

    [Fact]
    public void The_end_user_is_granted_execute_on_the_table_type()
    {
        var statements = AccessScript.Build(new AccessRequest
        {
            Name = "reader",
            ExistingUser = true,
            Scope = AccessScope.Procedure,
            Procedure = "[dbo].[p]",
            TableTypes = new[] { "[dbo].[team_list]", "[dbo].[team_list]" },
        });

        Assert.Equal(
            new[] { "GRANT EXECUTE ON OBJECT::[dbo].[p] TO [reader];", "GRANT EXECUTE ON TYPE::[dbo].[team_list] TO [reader];" },
            statements);
    }

    // ---------------------------------------------------------------- reports

    [Fact]
    public void A_one_column_table_is_a_multi_value_report_parameter_passed_as_json()
    {
        var ids = new SprocParameter
        {
            Name = "@Ids",
            SqlType = "id_list",
            TableType = "[dbo].[id_list]",
            Columns = new[] { Column("id", "int", nullable: false) },
        };

        var rdl = XDocument.Parse(SsrsClient.BuildStoredProcedureReportRdl(
            "[dbo].[p]", new[] { ids }, new[] { new ReportColumn { Name = "a" } }, RdlDataSource.Shared("/Data Sources/Main"),
            settings: ParameterSettings.FromProcedure(new[] { ids })));

        var parameter = rdl.Descendants(Rdl + "ReportParameter").Single();
        Assert.Equal("Ids", (string)parameter.Attribute("Name")!);
        Assert.Equal("Integer", (string)parameter.Element(Rdl + "DataType")!);
        Assert.Equal("true", (string)parameter.Element(Rdl + "MultiValue")!);
        Assert.Equal("false", (string)parameter.Element(Rdl + "Nullable")!);

        // Called from a batch, with the values as JSON.
        Assert.Empty(rdl.Descendants(Rdl + "CommandType"));
        Assert.Contains("OPENJSON(@Ids)", (string)rdl.Descendants(Rdl + "CommandText").First());
        Assert.Equal("=Code.FlankTable(Parameters!Ids.Value)", (string)rdl.Descendants(Rdl + "QueryParameter").Single().Element(Rdl + "Value")!);
        Assert.Contains("Public Function FlankTable", (string)rdl.Root!.Element(Rdl + "Code")!);
    }

    // ---------------------------------------------------------------- workbooks

    private static WorkbookParameter TeamsCell(WorkbookTableLayout layout, string label = "Teams") =>
        new WorkbookParameter("Teams", label)
        {
            Required = true,
            Note = "One team per row",
            Table = new WorkbookTable(new[]
            {
                new WorkbookParameter("team_id", "team_id") { Kind = WorkbookParameterKind.WholeNumber, Minimum = -2147483648, Maximum = 2147483647, Required = true },
                new WorkbookParameter("team_name", "team_name") { Kind = WorkbookParameterKind.Text, MaxLength = 50 },
            })
            {
                Layout = layout,
                Rows = 3,
            },
        };

    private WorkbookRequest TableRequest(WorkbookTableLayout layout) =>
        new WorkbookRequest("EXEC [dbo].[p] @League_Id = @League_Id, @Teams = @Teams, @Since = @Since", "my-server", "my-db")
        {
            Parameters = new[]
            {
                new WorkbookParameter("League_Id", "League") { Kind = WorkbookParameterKind.WholeNumber },
                TeamsCell(layout),
                new WorkbookParameter("Since", "Since") { Kind = WorkbookParameterKind.Date },
            },
            ParameterValues = new object?[] { 2, new List<IReadOnlyList<object?>> { new object?[] { 7, "Lakers" } }, null },
        };

    private XDocument ReadPart(string path)
    {
        using var zip = ZipFile.OpenRead(outputPath);
        using var stream = zip.GetEntry(path)!.Open();
        return XDocument.Load(stream);
    }

    private string Text(XDocument sheet, string reference) =>
        sheet.Descendants(Sheet + "c").SingleOrDefault(c => (string)c.Attribute("r")! == reference)?
            .Descendants(Sheet + "t").SingleOrDefault()?.Value ?? "";

    private Dictionary<string, string> Names() =>
        ReadPart("xl/workbook.xml").Descendants(Sheet + "definedName").ToDictionary(n => (string)n.Attribute("name")!, n => n.Value);

    [Fact]
    public void An_inline_table_sits_between_the_parameters_with_a_blank_row_after()
    {
        ExcelExporter.Generate(outputPath, TableRequest(WorkbookTableLayout.Inline));
        var sheet = ReadPart("xl/worksheets/sheet1.xml");

        // League (5); Teams' label and its table's header (6), rows (7-9), a hint (10);
        // Since (11); blank (12); query table (13). Column A only has names.
        Assert.Equal("League", Text(sheet, "A5"));
        Assert.Equal("Teams", Text(sheet, "A6"));
        Assert.Equal("team_id", Text(sheet, "B6"));
        Assert.Equal("team_name", Text(sheet, "C6"));
        Assert.Equal("7", sheet.Descendants(Sheet + "c").Single(c => (string)c.Attribute("r")! == "B7").Element(Sheet + "v")!.Value);
        Assert.Equal("Lakers", Text(sheet, "C7"));
        Assert.Equal("", Text(sheet, "A7"));
        Assert.StartsWith("More rows:", Text(sheet, "B10"));
        Assert.Equal("Since", Text(sheet, "A11"));

        var table = ReadPart("xl/tables/table3.xml").Root!;
        Assert.Equal("Teams", (string)table.Attribute("name")!);
        Assert.Equal("B6:C9", (string)table.Attribute("ref")!);
        Assert.Equal(new[] { "team_id", "team_name" }, table.Descendants(Sheet + "tableColumn").Select(c => (string)c.Attribute("name")!));

        Assert.Equal("3", (string)sheet.Root!.Element(Sheet + "tableParts")!.Attribute("count")!);
        Assert.Contains(ReadPart("xl/worksheets/_rels/sheet1.xml.rels").Root!.Elements(), r => (string)r.Attribute("Target")! == "../tables/table3.xml");
        Assert.Contains("/xl/tables/table3.xml", ReadPart("[Content_Types].xml").ToString());

        var names = Names();
        Assert.Equal("Query1!$B$5", names["Flank_League_Id"]);
        Assert.Equal("Query1!$B$11", names["Flank_Since"]);
        Assert.False(names.ContainsKey("Flank_Teams"));
        Assert.Equal("A13:A14", (string)ReadPart("xl/tables/table1.xml").Root!.Attribute("ref")!);

        // The note goes on the table's label.
        Assert.Equal("A6", (string)ReadPart("xl/comments1.xml").Descendants(Sheet + "comment").Single().Attribute("ref")!);
    }

    [Fact]
    public void An_inline_table_starts_with_as_many_rows_as_its_starting_values_if_more()
    {
        var request = TableRequest(WorkbookTableLayout.Inline);
        request.Parameters![1].Table!.Rows = 1;
        request.ParameterValues = new object?[] { null, new List<IReadOnlyList<object?>> { new object?[] { 1, "a" }, new object?[] { 2, "b" } }, null };

        ExcelExporter.Generate(outputPath, request);

        Assert.Equal("B6:C8", (string)ReadPart("xl/tables/table3.xml").Root!.Attribute("ref")!);
        Assert.Equal("Query1!$B$10", Names()["Flank_Since"]);
    }

    [Fact]
    public void Table_columns_get_checks_and_required_columns_are_highlighted_in_filled_rows()
    {
        ExcelExporter.Generate(outputPath, TableRequest(WorkbookTableLayout.Inline));
        var sheet = ReadPart("xl/worksheets/sheet1.xml").Root!;

        var validations = sheet.Descendants(Sheet + "dataValidation").ToDictionary(v => (string)v.Attribute("sqref")!);
        Assert.Equal("whole", (string)validations["B7:B9"].Attribute("type")!);
        Assert.Equal("textLength", (string)validations["C7:C9"].Attribute("type")!);

        var format = sheet.Elements(Sheet + "conditionalFormatting").Single(f => (string)f.Attribute("sqref")! == "B7:B9");
        Assert.Equal("AND(COUNTA($B7:$C7)>0,LEN(TRIM(B7))=0)", (string)format.Descendants(Sheet + "formula").Single());
    }

    [Fact]
    public void The_query_sends_a_tables_filled_rows_as_json()
    {
        ExcelExporter.Generate(outputPath, TableRequest(WorkbookTableLayout.Inline));
        string m = MashupReader.ReadSection1(outputPath);

        Assert.Contains("Table.SelectColumns(Excel.CurrentWorkbook(){[Name=\"Teams\"]}[Content], {\"team_id\", \"team_name\"})", m);
        Assert.Contains("not List.MatchesAll(r, (x) => x = null or (x is text and Text.Trim(x) = \"\"))", m);
        Assert.Contains("Text.FromBinary(Json.FromValue(Checked))", m);
        Assert.Contains("\"Teams needs at least one row. Enter one, then refresh.\"", m);
        Assert.Contains("at & \", team_id)\" & \" is required: fill it in, or clear the row.\"", m);
        Assert.Contains("at & \", team_name)\" & \" can be at most 50 characters.\"", m);
        Assert.Contains("Record.FromList({Value0, Value1, Value2}, {\"League_Id\", \"Teams\", \"Since\"})", m);

        Assert.Contains("rows in the tables", Text(ReadPart("xl/worksheets/sheet1.xml"), "A1"));
    }

    [Fact]
    public void A_table_can_have_its_own_sheet_linked_from_the_parameter_list()
    {
        ExcelExporter.Generate(outputPath, TableRequest(WorkbookTableLayout.Sheet));
        var sheet = ReadPart("xl/worksheets/sheet1.xml");

        // One row each, as without tables.
        Assert.Equal("Teams", Text(sheet, "A6"));
        Assert.Equal("Enter rows on the Teams sheet", Text(sheet, "B6"));
        var link = sheet.Descendants(Sheet + "hyperlink").Single();
        Assert.Equal("B6", (string)link.Attribute("ref")!);
        Assert.Equal("'Teams'!A1", (string)link.Attribute("location")!);
        var order = sheet.Root!.Elements().Select(e => e.Name.LocalName).ToList();
        Assert.True(order.IndexOf("hyperlinks") < order.IndexOf("pageMargins"));
        Assert.Equal("Query1!$B$7", Names()["Flank_Since"]);

        var book = ReadPart("xl/workbook.xml").Root!;
        Assert.Equal(new[] { "Query1", "Teams" }, book.Descendants(Sheet + "sheet").Select(s => (string)s.Attribute("name")!));

        var own = ReadPart("xl/worksheets/sheet3.xml");
        Assert.Equal("team_id", Text(own, "A1"));
        Assert.Equal("Lakers", Text(own, "B2"));
        Assert.Contains(own.Descendants(Sheet + "dataValidation"), v => (string)v.Attribute("sqref")! == "A2:A2");

        var table = ReadPart("xl/tables/table3.xml").Root!;
        Assert.Equal("A1:B2", (string)table.Attribute("ref")!);
        Assert.Equal("../tables/table3.xml", (string)ReadPart("xl/worksheets/_rels/sheet3.xml.rels").Root!.Elements().Single().Attribute("Target")!);

        var types = ReadPart("[Content_Types].xml").ToString();
        Assert.Contains("/xl/worksheets/sheet3.xml", types);
        Assert.Contains("/xl/tables/table3.xml", types);
    }

    [Fact]
    public void Own_sheets_come_after_the_hidden_lists_sheet_whose_names_refer_to_it_by_position()
    {
        var request = TableRequest(WorkbookTableLayout.Sheet);
        request.Parameters![0].List = new WorkbookList(null, "id", "id") { Rows = new (object?, object?)[] { (1, 1) } };

        ExcelExporter.Generate(outputPath, request);

        var book = ReadPart("xl/workbook.xml").Root!;
        Assert.Equal(new[] { "Query1", "DropdownOptions", "Teams" }, book.Descendants(Sheet + "sheet").Select(s => (string)s.Attribute("name")!));
        Assert.Equal("Teams", (string)ReadPart("xl/tables/table4.xml").Root!.Attribute("name")!);
    }

    [Fact]
    public void Table_and_sheet_names_are_valid_and_unique()
    {
        ExcelExporter.Generate(outputPath, new WorkbookRequest("EXEC p @a = @a, @b = @b", "s", "d")
        {
            Parameters = new[]
            {
                TeamsCell(WorkbookTableLayout.Sheet, "AB12"),
                TeamsCell(WorkbookTableLayout.Sheet, "AB12"),
            },
        });

        // AB12 is a cell, so not a table name; sheets can be named anything but must differ.
        Assert.Equal("Table_AB12", (string)ReadPart("xl/tables/table3.xml").Root!.Attribute("name")!);
        Assert.Equal("Table_AB12_2", (string)ReadPart("xl/tables/table4.xml").Root!.Attribute("name")!);
        var book = ReadPart("xl/workbook.xml").Root!;
        Assert.Equal(new[] { "Query1", "AB12", "AB12 (2)" }, book.Descendants(Sheet + "sheet").Select(s => (string)s.Attribute("name")!));
    }
}
