using Flank.Ssrs;
using System.Xml.Linq;

namespace Flank.Tests.Unit;

/// <summary>
/// Parameter defaults read from procedure definitions, and how the review
/// page's settings shape the procedure call and the report definition.
/// </summary>
public class ParameterSettingsTests
{
    private static readonly XNamespace Rdl =
        "http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition";

    private static SprocParameter Param(string name, string type, bool output = false, string? declaredDefault = null, bool hasDefault = false) =>
        new SprocParameter { Name = name, SqlType = type, IsOutput = output, MaxLength = 100, HasDefault = hasDefault, DefaultValue = declaredDefault };

    [Fact]
    public void Defaults_are_read_from_the_definition()
    {
        var defaults = ProcedureDefaults.Parse(@"
CREATE PROCEDURE dbo.p
    @required int,
    @count int = 5,
    @name nvarchar(50) = N'it''s',
    @since date = '2024-01-01',
    @maybe int = NULL,
    @negative decimal(5, 2) = -1.5,
    @flag bit = 1,
    @computed datetime = GETDATE   -- not a literal: has a default, but no value to show
AS SELECT 1;");

        Assert.False(defaults.ContainsKey("@required"));
        Assert.Equal("5", defaults["@count"].Value);
        Assert.Equal("it's", defaults["@name"].Value);
        Assert.Equal("2024-01-01", defaults["@since"].Value);
        Assert.Null(defaults["@maybe"].Value);
        Assert.Equal("-1.5", defaults["@negative"].Value);
        Assert.Equal("1", defaults["@flag"].Value);
        Assert.True(defaults.ContainsKey("@COUNT"));

        // As written, for "Omit (defaults to ...)".
        Assert.Equal("N'it''s'", defaults["@name"].Text);
        Assert.Equal("NULL", defaults["@maybe"].Text);
        Assert.Equal("-1.5", defaults["@negative"].Text);
        Assert.Equal("GETDATE", defaults["@computed"].Text);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("this isn't T-SQL")]
    public void Unreadable_definitions_give_no_defaults(string? definition)
    {
        Assert.Empty(ProcedureDefaults.Parse(definition));
    }

    [Fact]
    public void Settings_start_from_the_procedure()
    {
        var settings = ParameterSettings.FromProcedure(new[]
        {
            Param("@League_Id", "int"),
            Param("@Name", "nvarchar", declaredDefault: "x", hasDefault: true),
            Param("@Out", "int", output: true),
        });

        Assert.Equal(new[] { "League_Id", "Name" }, settings.Select(s => s.DisplayName));
        Assert.Equal(BlankValue.Required, settings[0].IfBlank);
        Assert.Equal(BlankValue.UseProcedureDefault, settings[1].IfBlank);
        Assert.Equal("x", settings[1].DefaultValue);
    }

    [Fact]
    public void Without_omitted_parameters_the_call_is_a_plain_exec()
    {
        var parameters = new[] { Param("@a", "int"), Param("@b", "int") };
        var settings = ParameterSettings.FromProcedure(parameters);
        settings[1].IfBlank = BlankValue.PassNull;

        Assert.Equal("EXEC [dbo].[p] @a = @a, @b = @b", ProcedureCall.BuildExec("[dbo].[p]", parameters, settings));
    }

    [Fact]
    public void Parameters_using_the_procedure_default_are_left_out_when_null()
    {
        var parameters = new[] { Param("@a", "int"), Param("@b", "nvarchar"), Param("@o", "int", output: true) };
        var settings = ParameterSettings.FromProcedure(parameters);
        settings[1].IfBlank = BlankValue.UseProcedureDefault;

        string call = ProcedureCall.BuildExec("[dbo].[it's]", parameters, settings);

        Assert.Contains("N', @a = @a'", call);
        Assert.Contains("CASE WHEN @b IS NOT NULL THEN N', @b = @b' END", call);
        Assert.Contains("N', @o = @o OUTPUT'", call);
        Assert.Contains("N'DECLARE @o int; EXEC [dbo].[it''s]'", call);
        Assert.Contains("EXEC sp_executesql @flank_call, N'@a int, @b nvarchar(50)', @a = @a, @b = @b;", call);
    }

    [Fact]
    public void Report_parameters_use_the_display_name_default_and_blank_setting()
    {
        var parameters = new[] { Param("@League_Id", "int"), Param("@Name", "nvarchar"), Param("@Max", "int") };
        var settings = ParameterSettings.FromProcedure(parameters);
        settings[0].DisplayName = "League";
        settings[0].DefaultValue = "2";
        settings[1].IfBlank = BlankValue.UseProcedureDefault;
        settings[2].IfBlank = BlankValue.PassNull;

        var rdl = XDocument.Parse(SsrsClient.BuildStoredProcedureReportRdl(
            "[dbo].[p]", parameters, new[] { new ReportColumn { Name = "x" } },
            RdlDataSource.AzureSql("s.database.windows.net", "d"), null, settings));

        var reportParameters = rdl.Descendants(Rdl + "ReportParameter").ToList();

        Assert.Equal("League", (string)reportParameters[0].Element(Rdl + "Prompt")!);
        Assert.Equal("false", (string)reportParameters[0].Element(Rdl + "Nullable")!);
        Assert.Equal("2", (string)reportParameters[0].Descendants(Rdl + "Value").Single());

        // Optional ones start blank (NULL) so the report opens with data.
        Assert.Equal("true", (string)reportParameters[1].Element(Rdl + "Nullable")!);
        Assert.Equal("=Nothing", (string)reportParameters[1].Descendants(Rdl + "Value").Single());
        Assert.Equal("=Nothing", (string)reportParameters[2].Descendants(Rdl + "Value").Single());

        // Leaving @Name out needs the built call, so the query is text, not a procedure.
        Assert.Empty(rdl.Descendants(Rdl + "CommandType"));
        Assert.StartsWith("DECLARE @flank_arguments", (string)rdl.Descendants(Rdl + "CommandText").Single());
    }

    [Fact]
    public void A_parameter_with_a_list_gets_a_dataset_and_valid_values()
    {
        var parameters = new[] { Param("@League_Id", "int") };
        var settings = ParameterSettings.FromProcedure(parameters);
        settings[0].ListProcedure = "[dbo].[get_leagues]";
        settings[0].ListValueColumn = "id";
        settings[0].ListLabelColumn = "name";

        var rdl = XDocument.Parse(SsrsClient.BuildStoredProcedureReportRdl(
            "[dbo].[p]", parameters, new[] { new ReportColumn { Name = "x" } },
            RdlDataSource.AzureSql("s.database.windows.net", "d"), null, settings));

        var list = rdl.Descendants(Rdl + "DataSet").Single(d => (string)d.Attribute("Name")! == "List_League_Id");
        Assert.Equal("StoredProcedure", (string)list.Descendants(Rdl + "CommandType").Single());
        Assert.Equal("[dbo].[get_leagues]", (string)list.Descendants(Rdl + "CommandText").Single());
        Assert.Equal("Database", (string)list.Descendants(Rdl + "DataSourceName").Single());
        Assert.Equal(new[] { "id", "name" }, list.Descendants(Rdl + "DataField").Select(f => (string)f));

        var reference = rdl.Descendants(Rdl + "ReportParameter").Single().Descendants(Rdl + "DataSetReference").Single();
        Assert.Equal("List_League_Id", (string)reference.Element(Rdl + "DataSetName")!);
        Assert.Equal("Value", (string)reference.Element(Rdl + "ValueField")!);
        Assert.Equal("Label", (string)reference.Element(Rdl + "LabelField")!);
    }

    [Fact]
    public void Without_settings_reports_are_as_before()
    {
        var parameters = new[] { Param("@League_Id", "int") };

        var rdl = XDocument.Parse(SsrsClient.BuildStoredProcedureReportRdl(
            "[dbo].[p]", parameters, new[] { new ReportColumn { Name = "x" } }, "/Data Sources/Main"));

        var reportParameter = rdl.Descendants(Rdl + "ReportParameter").Single();
        Assert.Equal("League_Id", (string)reportParameter.Element(Rdl + "Prompt")!);
        Assert.Equal("true", (string)reportParameter.Element(Rdl + "Nullable")!);
        Assert.Empty(reportParameter.Descendants(Rdl + "DefaultValue"));
        Assert.Equal("StoredProcedure", (string)rdl.Descendants(Rdl + "CommandType").Single());
    }
}
