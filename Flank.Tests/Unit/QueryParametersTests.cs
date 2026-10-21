using Flank.Ssrs;
using System.Xml.Linq;

namespace Flank.Tests.Unit;

public class QueryParametersTests
{
    private static readonly XNamespace Rdl =
        "http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition";

    [Fact]
    public void Finds_undeclared_variables_in_order_of_first_use()
    {
        var names = QueryParameters.Find(
            "SELECT * FROM teams WHERE league_id = @League_Id AND founded >= @since AND league_id <> @league_id");

        Assert.Equal(new[] { "@League_Id", "@since" }, names);
    }

    [Fact]
    public void Declared_variables_and_table_variables_are_not_parameters()
    {
        var names = QueryParameters.Find(
            "DECLARE @cutoff date = DATEADD(year, -1, @today);\r\n" +
            "DECLARE @t TABLE (id int);\r\n" +
            "INSERT @t SELECT id FROM teams WHERE founded >= @cutoff;\r\n" +
            "SELECT * FROM @t WHERE id > @min_id");

        Assert.Equal(new[] { "@today", "@min_id" }, names);
    }

    [Fact]
    public void Global_variables_and_strings_are_not_parameters()
    {
        Assert.Empty(QueryParameters.Find("SELECT @@ROWCOUNT, '@not_a_parameter' -- @nor_this"));
    }

    [Fact]
    public void A_query_that_does_not_parse_has_no_parameters_and_a_problem()
    {
        Assert.Empty(QueryParameters.Find("SELECT FROM WHERE @x"));
        Assert.StartsWith("Line 1:", QueryParameters.ParseProblem("SELECT FROM WHERE @x"));
        Assert.Null(QueryParameters.ParseProblem("SELECT @x"));
    }

    [Theory]
    [InlineData("int", "int", 0, 0, 0)]
    [InlineData("nvarchar(50)", "nvarchar", 100, 0, 0)]
    [InlineData("NVARCHAR(MAX)", "nvarchar", -1, 0, 0)]
    [InlineData("varchar(20)", "varchar", 20, 0, 0)]
    [InlineData("decimal(10, 2)", "decimal", 0, 10, 2)]
    [InlineData("decimal", "decimal", 0, 18, 0)]
    [InlineData("datetime2", "datetime2", 0, 0, 7)]
    [InlineData(" date ", "date", 0, 0, 0)]
    public void Parses_types_as_sys_parameters_would_describe_them(string text, string type, int maxLength, int precision, int scale)
    {
        var parameter = QueryParameters.Parameter("since", text)!;

        Assert.Equal("@since", parameter.Name);
        Assert.Equal(type, parameter.SqlType);
        Assert.Equal(maxLength, parameter.MaxLength);
        Assert.Equal(precision, parameter.Precision);
        Assert.Equal(scale, parameter.Scale);
        Assert.False(parameter.HasDefault);
    }

    [Theory]
    [InlineData("int(4)")]
    [InlineData("nvarchar(")]
    [InlineData("geography")]
    [InlineData("")]
    public void Rejects_unknown_or_malformed_types(string text)
    {
        Assert.Null(QueryParameters.Parameter("x", text));
    }

    [Fact]
    public void Types_round_trip_to_declarations()
    {
        // As declared on the procedure or in sp_executesql, and parsed back.
        foreach (string type in QueryParameters.CommonTypes)
        {
            string declared = ProcedureCall.TypeDeclaration(QueryParameters.Parameter("x", type)!);

            Assert.Equal(type.Replace("datetime2", "datetime2(7)"), declared);
            Assert.Equal(declared, ProcedureCall.TypeDeclaration(QueryParameters.Parameter("x", declared)!));
        }
    }

    [Fact]
    public void Procedure_script_declares_the_parameters_and_wraps_the_query()
    {
        string script = QueryParameters.CreateProcedureScript(
            "[dbo].[teams_since]",
            new[] { QueryParameters.Parameter("League_Id", "int")!, QueryParameters.Parameter("since", "date")! },
            "SELECT * FROM teams WHERE league_id = @League_Id AND founded >= @since;  ");

        Assert.Equal(
            "CREATE PROCEDURE [dbo].[teams_since]\n    @League_Id int,\n    @since date\nAS\nBEGIN\n    SET NOCOUNT ON;\n\n" +
            "SELECT * FROM teams WHERE league_id = @League_Id AND founded >= @since;\nEND;",
            script);
    }

    [Fact]
    public void Text_report_passes_the_parameters_to_the_query()
    {
        var parameters = new[] { QueryParameters.Parameter("League_Id", "int")!, QueryParameters.Parameter("since", "date")! };
        var settings = ParameterSettings.FromProcedure(parameters);
        settings[0].DisplayName = "League";

        var rdl = XDocument.Parse(SsrsClient.BuildTextReportRdl(
            "SELECT team_name FROM teams WHERE league_id = @League_Id AND founded >= @since",
            new[] { new ReportColumn { Name = "team_name" } },
            RdlDataSource.Shared("/Data Sources/Main"),
            parameters,
            settings));

        Assert.DoesNotContain(rdl.Descendants(Rdl + "CommandType"), c => (string)c == "StoredProcedure");
        Assert.Contains("@League_Id", (string)rdl.Descendants(Rdl + "CommandText").Single());

        var queryParams = rdl.Descendants(Rdl + "QueryParameter").ToList();
        Assert.Equal(new[] { "@League_Id", "@since" }, queryParams.Select(p => (string)p.Attribute("Name")!));

        var reportParams = rdl.Descendants(Rdl + "ReportParameter").ToList();
        Assert.Equal(new[] { "Integer", "DateTime" }, reportParams.Select(p => (string)p.Element(Rdl + "DataType")!));
        Assert.Equal("League", (string)reportParams[0].Element(Rdl + "Prompt")!);
    }

    [Fact]
    public void Text_report_without_parameters_is_the_plain_text_report()
    {
        string sql = "SELECT id FROM leagues";
        var columns = new[] { new ReportColumn { Name = "id" } };

        Assert.Equal(
            SsrsClient.BuildTextReportRdl(sql, columns, "/Data Sources/Main"),
            SsrsClient.BuildTextReportRdl(sql, columns, RdlDataSource.Shared("/Data Sources/Main"), Array.Empty<SprocParameter>(), null));
    }

    [Fact]
    public void Turning_a_query_into_a_procedure_for_access_declares_its_parameters()
    {
        var script = AccessScript.Build(new AccessRequest
        {
            SignIn = AccessSignIn.SqlLogin,
            Name = "report_user",
            Password = "x",
            Database = "sales",
            IsAzureSqlDatabase = true,
            Scope = AccessScope.Procedure,
            Procedure = "[dbo].[flank_report]",
            ProcedureBody = "SELECT id FROM orders WHERE placed >= @since",
            ProcedureParameters = new[] { QueryParameters.Parameter("since", "date")! },
        });

        Assert.Contains(script, s => s.StartsWith("CREATE PROCEDURE [dbo].[flank_report]\n    @since date\nAS"));
    }
}
