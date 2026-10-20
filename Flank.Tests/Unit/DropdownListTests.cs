using Flank.Excel;
using Flank.Ssrs;
using Flank.Tests.Support;
using System.Xml.Linq;

namespace Flank.Tests.Unit;

/// <summary>Dropdown choices from a query or typed-in values, as reports and workbooks hold them.</summary>
public class DropdownListTests : IDisposable
{
    private static readonly XNamespace Rdl =
        "http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition";

    private readonly string outputPath =
        Path.Combine(Path.GetTempPath(), $"flank-unit-{Guid.NewGuid():N}.xlsx");

    public void Dispose()
    {
        File.Delete(outputPath);
    }

    private static SprocParameter Param(string type, short maxLength = 0) =>
        new SprocParameter { Name = "@p", SqlType = type, MaxLength = maxLength };

    private static IReadOnlyList<KeyValuePair<object?, object?>> Parse(SprocParameter parameter, string values, string? labels = null)
    {
        var rows = ProcedureList.ParseValues(parameter, values, labels, out string? problem);
        Assert.True(rows != null, problem);
        return rows!;
    }

    private static string Problem(SprocParameter parameter, string values, string? labels = null)
    {
        Assert.Null(ProcedureList.ParseValues(parameter, values, labels, out string? problem));
        return problem!;
    }

    [Fact]
    public void Values_only_are_their_own_labels_one_per_line()
    {
        var rows = Parse(Param("int"), "1\r\n\r\n 2 \n3\n");

        Assert.Equal(new object?[] { 1L, 2L, 3L }, rows.Select(r => r.Key));
        Assert.Equal(new object?[] { "1", "2", "3" }, rows.Select(r => r.Value));
    }

    [Fact]
    public void Values_with_labels_pair_up_by_line()
    {
        var rows = Parse(Param("int"), "1\n2", "NBA\nNFL");

        Assert.Equal(new object?[] { "NBA", "NFL" }, rows.Select(r => r.Value));
    }

    [Fact]
    public void Values_and_labels_must_be_the_same_length()
    {
        Assert.Contains("2 value(s) but 3 label(s)", Problem(Param("int"), "1\n2", "NBA\nNFL\nMLB"));
    }

    [Theory]
    [InlineData("int", "1\nabc", "Value 2, \"abc\": not a whole number.")]
    [InlineData("tinyint", "300", "out of range for tinyint")]
    [InlineData("bit", "yes", "not 1, 0, true or false.")]
    [InlineData("date", "someday", "not a date.")]
    [InlineData("decimal", "1.5\nx", "not a number.")]
    public void Values_must_be_of_the_parameter_type(string type, string values, string expected)
    {
        Assert.Contains(expected, Problem(Param(type), values));
    }

    [Fact]
    public void Text_values_must_fit_the_length()
    {
        Assert.Contains("longer than nvarchar(3)", Problem(Param("nvarchar", maxLength: 6), "abcd"));
        Assert.Single(Parse(Param("nvarchar", maxLength: 6), "abc"));
    }

    [Fact]
    public void Values_and_labels_must_not_repeat()
    {
        Assert.Contains("in the values twice", Problem(Param("int"), "1\n1"));
        Assert.Contains("in the labels twice", Problem(Param("int"), "1\n2", "A\na"));
    }

    [Fact]
    public void Dates_and_decimals_read_the_same_on_any_windows_settings()
    {
        var rows = Parse(Param("date"), "2024-01-31");
        Assert.Equal(new DateTime(2024, 1, 31), rows[0].Key);
        Assert.Equal("2024-01-31", ProcedureList.FormatValue(rows[0].Key));

        Assert.Equal(1.5m, Parse(Param("decimal"), "1.5")[0].Key);
    }

    [Fact]
    public void Nothing_typed_is_a_problem()
    {
        Assert.Contains("at least one value", Problem(Param("int"), " \n "));
    }

    private static IReadOnlyList<ParameterSettings> Settings(Action<ParameterSettings> setUp)
    {
        var settings = ParameterSettings.FromProcedure(new[] { new SprocParameter { Name = "@League_Id", SqlType = "int" } });
        setUp(settings[0]);
        return settings;
    }

    private static XDocument Report(IReadOnlyList<ParameterSettings> settings) =>
        XDocument.Parse(SsrsClient.BuildStoredProcedureReportRdl(
            "[dbo].[p]", settings.Select(s => s.Parameter).ToList(), new[] { new ReportColumn { Name = "x" } },
            RdlDataSource.AzureSql("s.database.windows.net", "d"), null, settings));

    [Fact]
    public void Typed_values_are_written_into_the_report_parameter()
    {
        var settings = Settings(s => ProcedureList.ApplyValues(
            s, Parse(s.Parameter, "1\n2", "NBA & co\nNFL"), withLabels: true));

        var rdl = Report(settings);

        Assert.DoesNotContain(rdl.Descendants(Rdl + "DataSet"), d => ((string)d.Attribute("Name")!).StartsWith("List_"));
        var values = rdl.Descendants(Rdl + "ParameterValue").ToList();
        Assert.Equal(new[] { "1", "2" }, values.Select(v => (string)v.Element(Rdl + "Value")!));
        Assert.Equal(new[] { "NBA & co", "NFL" }, values.Select(v => (string)v.Element(Rdl + "Label")!));
    }

    [Fact]
    public void A_query_list_is_a_text_dataset()
    {
        var settings = Settings(s =>
        {
            s.ListSource = ListSource.Query;
            s.ListQuery = "SELECT id, name FROM leagues WHERE id < 3";
            s.ListValueColumn = "id";
            s.ListLabelColumn = "name";
        });

        var list = Report(settings).Descendants(Rdl + "DataSet").Single(d => (string)d.Attribute("Name")! == "List_League_Id");

        Assert.Empty(list.Descendants(Rdl + "CommandType"));
        Assert.Equal("SELECT id, name FROM leagues WHERE id < 3", (string)list.Descendants(Rdl + "CommandText").Single());
        Assert.Equal(new[] { "id", "name" }, list.Descendants(Rdl + "DataField").Select(f => (string)f));
    }

    [Fact]
    public void List_sql_and_has_list_follow_the_source()
    {
        var settings = Settings(s => { s.ListProcedure = "[dbo].[get_leagues]"; s.ListValueColumn = "id"; s.ListLabelColumn = "name"; })[0];
        Assert.Equal("EXEC [dbo].[get_leagues]", settings.ListSql);
        Assert.True(settings.HasList);

        settings.ListSource = ListSource.Values;
        Assert.Null(settings.ListSql);
        Assert.False(settings.HasList); // no values yet

        ProcedureList.Clear(settings);
        Assert.False(settings.HasList);
        Assert.Equal(ListSource.Procedure, settings.ListSource);
    }

    [Fact]
    public void Excel_holds_typed_values_as_a_table_with_no_query()
    {
        ExcelExporter.Generate(outputPath, new WorkbookRequest("EXEC [dbo].[p] @League_Id = @League_Id, @Since = @Since", "my-server", "my-db")
        {
            Parameters = new[]
            {
                new WorkbookParameter("League_Id", "League")
                {
                    List = new WorkbookList(null, ProcedureList.ValueColumn, ProcedureList.LabelColumn)
                    {
                        Rows = new (object?, object?)[] { (1L, "NBA \"pro\""), (2L, "NFL") },
                    },
                },
                new WorkbookParameter("Since", "Since")
                {
                    List = new WorkbookList(null, ProcedureList.ValueColumn, ProcedureList.ValueColumn)
                    {
                        Rows = new (object?, object?)[] { (new DateTime(2024, 1, 31), "2024-01-31") },
                    },
                },
            },
            ParameterValues = new object?[] { 2L, null },
        });

        string m = MashupReader.ReadSection1(outputPath);

        Assert.Contains("shared League_options = #table({\"label\", \"value\"}, {{\"NBA \"\"pro\"\"\", 1}, {\"NFL\", 2}});", m);
        Assert.Contains("shared Since_options = #table({\"value\"}, {{#date(2024, 1, 31)}});", m);
        Assert.DoesNotContain("Sql.Database(\"my-server\", \"my-db\", [Query=", m);
    }
}
