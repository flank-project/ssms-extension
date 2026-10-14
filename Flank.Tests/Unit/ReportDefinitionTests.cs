using Flank.Ssrs;
using System.Xml.Linq;

namespace Flank.Tests.Unit;

public class ReportDefinitionTests
{
    private static readonly XNamespace Rdl =
        "http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition";

    private static IReadOnlyList<ReportColumn> Columns(params string?[] names) =>
        names.Select(n => new ReportColumn { Name = n }).ToList();

    private static SprocParameter Param(string name, string type, bool output = false) =>
        new SprocParameter { Name = name, SqlType = type, IsOutput = output };

    private static XDocument TextReport(string sql, params string?[] columns) =>
        XDocument.Parse(SsrsClient.BuildTextReportRdl(sql, Columns(columns), "/Data Sources/Main"));

    private static XDocument SprocReport(IReadOnlyList<SprocParameter> parameters, params string?[] columns) =>
        XDocument.Parse(SsrsClient.BuildStoredProcedureReportRdl(
            "dbo.get_things", parameters, Columns(columns), "/Data Sources/Main"));

    private static List<string> FieldNames(XDocument rdl) =>
        rdl.Descendants(Rdl + "Field").Select(f => (string)f.Attribute("Name")!).ToList();

    [Fact]
    public void Text_report_uses_the_sql_and_shared_data_source()
    {
        var rdl = TextReport("SELECT a, b FROM t WHERE a < 3 AND b <> 'x & y'", "a", "b");

        Assert.Equal(
            "SELECT a, b FROM t WHERE a < 3 AND b <> 'x & y'",
            (string)rdl.Descendants(Rdl + "CommandText").Single());
        Assert.Equal("/Data Sources/Main", (string)rdl.Descendants(Rdl + "DataSourceReference").Single());
        Assert.Empty(rdl.Descendants(Rdl + "CommandType"));
        Assert.Empty(rdl.Descendants(Rdl + "ReportParameters"));
    }

    [Fact]
    public void Each_column_gets_a_field_header_and_value()
    {
        var rdl = TextReport("SELECT 1", "id", "name");

        Assert.Equal(new[] { "id", "name" }, FieldNames(rdl));
        Assert.Equal(
            new[] { "id", "name" },
            rdl.Descendants(Rdl + "DataField").Select(d => (string)d));
        Assert.Equal(2, rdl.Descendants(Rdl + "TablixColumn").Count());

        var values = rdl.Descendants(Rdl + "Value").Select(v => (string)v).ToList();
        Assert.Contains("=Fields!id.Value", values);
        Assert.Contains("=Fields!name.Value", values);
    }

    [Fact]
    public void Column_names_are_made_safe_for_rdl()
    {
        var rdl = TextReport("SELECT 1", "order id", "1st", "total$", null);

        Assert.Equal(new[] { "order_id", "_1st", "total_", "Column4" }, FieldNames(rdl));
        // DataField still points at the real column name.
        Assert.Equal(
            new[] { "order id", "1st", "total$", "Column4" },
            rdl.Descendants(Rdl + "DataField").Select(d => (string)d));
    }

    [Fact]
    public void Field_names_are_unique_when_column_names_collide()
    {
        // SSRS rejects a report with duplicate field or textbox names.
        var rdl = TextReport("SELECT 1", "order id", "order_id", "x", "x");

        var names = FieldNames(rdl);
        Assert.Equal(names.Count, names.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    [Fact]
    public void Sproc_report_calls_the_procedure_with_its_parameters()
    {
        var rdl = SprocReport(
            new[] { Param("@League_Id", "int"), Param("@Name", "varchar") },
            "team_name");

        Assert.Equal("StoredProcedure", (string)rdl.Descendants(Rdl + "CommandType").Single());
        Assert.Equal("dbo.get_things", (string)rdl.Descendants(Rdl + "CommandText").Single());

        var queryParams = rdl.Descendants(Rdl + "QueryParameter").ToList();
        Assert.Equal(new[] { "@League_Id", "@Name" }, queryParams.Select(p => (string)p.Attribute("Name")!));
        Assert.Equal("=Parameters!League_Id.Value", (string)queryParams[0].Element(Rdl + "Value")!);

        var reportParams = rdl.Descendants(Rdl + "ReportParameter").ToList();
        Assert.Equal(new[] { "League_Id", "Name" }, reportParams.Select(p => (string)p.Attribute("Name")!));
        Assert.Equal(new[] { "Integer", "String" }, reportParams.Select(p => (string)p.Element(Rdl + "DataType")!));
    }

    [Fact]
    public void Sproc_report_parameters_get_default_values()
    {
        var rdl = XDocument.Parse(SsrsClient.BuildStoredProcedureReportRdl(
            "dbo.get_things",
            new[] { Param("@Id", "int"), Param("@From", "date"), Param("@Active", "bit"), Param("@Name", "varchar") },
            Columns("x"),
            "/Data Sources/Main",
            new Dictionary<string, object?>
            {
                ["@Id"] = 2L,
                ["@From"] = new DateTime(2024, 1, 2),
                ["@Active"] = true,
                ["@Name"] = null,
            }));

        string? Default(string name) =>
            rdl.Descendants(Rdl + "ReportParameter")
                .Single(p => (string)p.Attribute("Name")! == name)
                .Descendants(Rdl + "Value").Select(v => (string)v).SingleOrDefault();

        Assert.Equal("2", Default("Id"));
        Assert.Equal("2024-01-02T00:00:00", Default("From"));
        Assert.Equal("true", Default("Active"));
        Assert.Null(Default("Name"));
    }

    [Fact]
    public void Sproc_report_skips_output_parameters()
    {
        var rdl = SprocReport(
            new[] { Param("@InputValue", "int"), Param("@Result", "int", output: true) },
            "x");

        Assert.Equal(
            new[] { "InputValue" },
            rdl.Descendants(Rdl + "ReportParameter").Select(p => (string)p.Attribute("Name")!));
    }

    [Fact]
    public void Sproc_report_without_parameters_has_no_parameter_sections()
    {
        var rdl = SprocReport(Array.Empty<SprocParameter>(), "id");

        Assert.Empty(rdl.Descendants(Rdl + "ReportParameters"));
        Assert.Empty(rdl.Descendants(Rdl + "QueryParameters"));
    }

    [Fact]
    public void Sproc_report_with_output_parameters_calls_it_from_a_batch()
    {
        var rdl = SprocReport(
            new[] { Param("@N", "int"), Param("@Total", "int", output: true) },
            "x");

        Assert.Empty(rdl.Descendants(Rdl + "CommandType"));
        Assert.Equal(
            "DECLARE @Total int; EXEC dbo.get_things @N = @N, @Total = @Total OUTPUT",
            (string)rdl.Descendants(Rdl + "CommandText").Single());
        Assert.Equal(
            new[] { "@N" },
            rdl.Descendants(Rdl + "QueryParameter").Select(p => (string)p.Attribute("Name")!));
    }

    [Fact]
    public void Exec_without_parameters_is_just_the_call()
    {
        Assert.Equal("EXEC [dbo].[p]", ProcedureCall.BuildExec("[dbo].[p]", Array.Empty<SprocParameter>()));
    }

    [Theory]
    [InlineData("int", 4, 10, 0, "int")]
    [InlineData("varchar", 50, 0, 0, "varchar(50)")]
    [InlineData("varchar", -1, 0, 0, "varchar(max)")]
    [InlineData("nvarchar", 100, 0, 0, "nvarchar(50)")]
    [InlineData("nvarchar", -1, 0, 0, "nvarchar(max)")]
    [InlineData("decimal", 9, 10, 2, "decimal(10, 2)")]
    [InlineData("datetime2", 8, 27, 7, "datetime2(7)")]
    public void Declares_output_variables_with_their_full_type(string type, short maxLength, byte precision, byte scale, string expected)
    {
        var parameter = new SprocParameter { Name = "@o", SqlType = type, MaxLength = maxLength, Precision = precision, Scale = scale, IsOutput = true };

        Assert.Equal(expected, ProcedureCall.TypeDeclaration(parameter));
    }

    [Theory]
    [InlineData("int", "Integer")]
    [InlineData("SMALLINT", "Integer")]
    [InlineData("bigint", "String")]
    [InlineData("decimal", "Float")]
    [InlineData("money", "Float")]
    [InlineData("bit", "Boolean")]
    [InlineData("date", "DateTime")]
    [InlineData("datetime2", "DateTime")]
    [InlineData("nvarchar", "String")]
    [InlineData("uniqueidentifier", "String")]
    public void Maps_sql_parameter_types_to_rdl_types(string sqlType, string rdlType)
    {
        Assert.Equal(rdlType, SsrsClient.GetRdlParameterType(sqlType));
    }

    [Fact]
    public void Rejects_unsupported_parameter_types()
    {
        Assert.Throws<NotSupportedException>(() => SsrsClient.GetRdlParameterType("geography"));
    }
}
