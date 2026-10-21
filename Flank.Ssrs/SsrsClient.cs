using Flank.Ssrs.ReportService2010;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace Flank.Ssrs
{
    public class SsrsException : Exception
    {
        public SsrsException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    public class SsrsClient
    {
        private readonly string _reportServerUrl;
        private readonly string _reportPortalUrl;
        private readonly ICredentials _credentials;

        public SsrsClient(
            string reportServerUrl,
            string reportPortalUrl,
            ICredentials credentials = null)
        {
            _reportServerUrl = reportServerUrl.TrimEnd('/');
            _reportPortalUrl = reportPortalUrl.TrimEnd('/');
            _credentials = credentials;
        }

        private class TableRdl
        {
            public string Fields { get; set; }
            public string Columns { get; set; }
            public string HeaderCells { get; set; }
            public string DetailCells { get; set; }
            public string ColumnMembers { get; set; }
        }

        public void TestConnection()
        {
            try
            {
                var ssrs = CreateSsrsConnection();
                ssrs.ListChildren("/", false);
            }
            catch (Exception ex)
            {
                throw new SsrsException(
                    $"Could not connect to SSRS at '{_reportServerUrl}'.",
                    ex);
            }
        }

        public string CreateReportFromText(
            string sql,
            IReadOnlyList<ReportColumn> columns,
            string sharedDataSourcePath,
            string reportFolder,
            string reportName)
        {
            string rdl = BuildTextReportRdl(
                sql,
                columns,
                sharedDataSourcePath);

            DeployReport(
                rdl,
                reportFolder,
                reportName);

            return GetReportUrl(
                reportFolder,
                reportName);
        }

        /// <summary>A report on a query with parameters; see <see cref="BuildTextReportRdl(string, IReadOnlyList{ReportColumn}, RdlDataSource, IReadOnlyList{SprocParameter}, IReadOnlyList{ParameterSettings}?)"/>.</summary>
        public string CreateReportFromText(
            string sql,
            IReadOnlyList<ReportColumn> columns,
            IReadOnlyList<SprocParameter> parameters,
            IReadOnlyList<ParameterSettings>? settings,
            string sharedDataSourcePath,
            string reportFolder,
            string reportName)
        {
            string rdl = BuildTextReportRdl(
                sql,
                columns,
                RdlDataSource.Shared(sharedDataSourcePath),
                parameters,
                settings);

            DeployReport(
                rdl,
                reportFolder,
                reportName);

            return GetReportUrl(
                reportFolder,
                reportName);
        }

        public string CreateReportFromStoredProcedure(
            string procedureName,
            IReadOnlyList<SprocParameter> parameters,
            IReadOnlyList<ReportColumn> columns,
            string sharedDataSourcePath,
            string reportFolder,
            string reportName,
            IReadOnlyDictionary<string, object?>? defaultValues = null,
            IReadOnlyList<ParameterSettings>? settings = null)
        {
            string rdl = BuildStoredProcedureReportRdl(
                procedureName,
                parameters,
                columns,
                RdlDataSource.Shared(sharedDataSourcePath),
                defaultValues,
                settings);

            DeployReport(
                rdl,
                reportFolder,
                reportName);

            return GetReportUrl(
                reportFolder,
                reportName);
        }
        // RDL building is separate from deployment so it can be tested without a server.
        internal static string BuildTextReportRdl(
            string sql,
            IReadOnlyList<ReportColumn> columns,
            string sharedDataSourcePath) =>
            BuildTextReportRdl(sql, columns, RdlDataSource.Shared(sharedDataSourcePath));

        /// <summary>
        /// A report on a query with parameters (@names it doesn't DECLARE): each
        /// becomes a report parameter, set up as the review page chose, and is
        /// passed to the query as a query parameter.
        /// </summary>
        public static string BuildTextReportRdl(
            string sql,
            IReadOnlyList<ReportColumn> columns,
            RdlDataSource dataSource,
            IReadOnlyList<SprocParameter> parameters,
            IReadOnlyList<ParameterSettings>? settings)
        {
            if (parameters.Count == 0)
                return BuildTextReportRdl(sql, columns, dataSource);

            BuildStoredProcedureParameters(
                parameters,
                null,
                settings,
                out string reportParametersXml,
                out string queryParametersXml);

            string queryXml = $@"
        <CommandText>{XmlEscape(sql)}</CommandText>
{queryParametersXml}";

            return BuildReportRdl(
                columns,
                dataSource,
                queryXml,
                reportParametersXml,
                ListDataSetsXml(settings, dataSource));
        }

        public static string BuildTextReportRdl(
            string sql,
            IReadOnlyList<ReportColumn> columns,
            RdlDataSource dataSource)
        {
            string queryXml = $@"
        <CommandText>{XmlEscape(sql)}</CommandText>";

            return BuildReportRdl(
                columns,
                dataSource,
                queryXml,
                "");
        }

        internal static string BuildStoredProcedureReportRdl(
            string procedureName,
            IReadOnlyList<SprocParameter> parameters,
            IReadOnlyList<ReportColumn> columns,
            string sharedDataSourcePath,
            IReadOnlyDictionary<string, object?>? defaultValues = null) =>
            BuildStoredProcedureReportRdl(
                procedureName,
                parameters,
                columns,
                RdlDataSource.Shared(sharedDataSourcePath),
                defaultValues);

        /// <param name="defaultValues">
        /// Report parameter defaults by SQL parameter name (with '@'), so the
        /// report shows data when opened. Missing or null: no default. Takes
        /// precedence over the settings' defaults.
        /// </param>
        /// <param name="settings">
        /// From the parameter review page: prompts, defaults and what a blank
        /// value means. Null: as the procedure declares them.
        /// </param>
        public static string BuildStoredProcedureReportRdl(
            string procedureName,
            IReadOnlyList<SprocParameter> parameters,
            IReadOnlyList<ReportColumn> columns,
            RdlDataSource dataSource,
            IReadOnlyDictionary<string, object?>? defaultValues = null,
            IReadOnlyList<ParameterSettings>? settings = null)
        {
            string reportParametersXml;
            string queryParametersXml;

            BuildStoredProcedureParameters(
                parameters,
                defaultValues,
                settings,
                out reportParametersXml,
                out queryParametersXml);

            // SSRS can't supply OUTPUT or table-valued parameters to a stored
            // procedure call, or leave a parameter out when it's blank, so then
            // the procedure is called from a batch.
            string call = ProcedureCall.BuildExec(procedureName, parameters, settings);
            bool needsBatch = parameters.Any(p => p.IsOutput || p.IsTable) || call.Contains("@flank_arguments");

            string queryXml = needsBatch
                ? $@"
        <CommandText>{XmlEscape(call)}</CommandText>
{queryParametersXml}"
                : $@"
        <CommandType>StoredProcedure</CommandType>
        <CommandText>{XmlEscape(procedureName)}</CommandText>
{queryParametersXml}";

            return BuildReportRdl(
                columns,
                dataSource,
                queryXml,
                reportParametersXml,
                ListDataSetsXml(settings, dataSource),
                parameters.Any(p => p.IsTable && !p.IsOutput) ? TableCodeXml : "");
        }

        /// <summary>
        /// Report code that turns a multi-value parameter's values into the
        /// JSON a one-column table-valued parameter is passed as (see
        /// <see cref="TableValues"/>): [["a"],["b"]].
        /// </summary>
        private const string TableCodeXml = @"
  <Code>Public Function FlankTable(ByVal values As Object) As String
  If values Is Nothing Then Return Nothing
  Dim q As String = ChrW(34)
  Dim json As New System.Text.StringBuilder(""["")
  Dim first As Boolean = True
  For Each v As Object In values
    If Not first Then json.Append("","")
    first = False
    json.Append(""["")
    If v Is Nothing Then
      json.Append(""null"")
    ElseIf TypeOf v Is String Then
      json.Append(q).Append(CStr(v).Replace(""\"", ""\\"").Replace(q, ""\"" &amp; q).Replace(vbCr, ""\r"").Replace(vbLf, ""\n"").Replace(vbTab, ""\t"")).Append(q)
    ElseIf TypeOf v Is Boolean Then
      json.Append(If(CBool(v), ""true"", ""false""))
    ElseIf TypeOf v Is DateTime Then
      json.Append(q).Append(CDate(v).ToString(""yyyy-MM-dd'T'HH:mm:ss"", System.Globalization.CultureInfo.InvariantCulture)).Append(q)
    Else
      json.Append(System.Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture))
    End If
    json.Append(""]"")
  Next
  Return json.Append(""]"").ToString()
End Function</Code>";

        private static string ListDataSetName(string reportParameterName) =>
            "List_" + MakeSafeRdlName(reportParameterName, 0);

        /// <summary>
        /// A dataset per parameter with a list procedure: the procedure called
        /// with no parameters, its value and label columns as fields Value and
        /// Label, for the parameter's valid values.
        /// </summary>
        private static string ListDataSetsXml(IReadOnlyList<ParameterSettings>? settings, RdlDataSource dataSource)
        {
            var dataSets = new StringBuilder();

            foreach (var setting in settings ?? new ParameterSettings[0])
            {
                // Typed-in values are written into the parameter itself.
                if (!setting.HasList || setting.ListSource == ListSource.Values)
                    continue;

                string commandXml = setting.ListSource == ListSource.Query
                    ? $@"
        <CommandText>{XmlEscape(setting.ListQuery!)}</CommandText>"
                    : $@"
        <CommandType>StoredProcedure</CommandType>
        <CommandText>{XmlEscape(setting.ListProcedure!)}</CommandText>";

                dataSets.Append($@"
    <DataSet Name=""{XmlEscape(ListDataSetName(setting.SqlName.TrimStart('@')))}"">
      <Query>
        <DataSourceName>{XmlEscape(dataSource.Name)}</DataSourceName>{commandXml}
      </Query>
      <Fields>
        <Field Name=""Value"">
          <DataField>{XmlEscape(setting.ListValueColumn!)}</DataField>
        </Field>
        <Field Name=""Label"">
          <DataField>{XmlEscape(setting.ListLabelColumn!)}</DataField>
        </Field>
      </Fields>
    </DataSet>");
            }

            return dataSets.ToString();
        }

        private static void BuildStoredProcedureParameters(
            IReadOnlyList<SprocParameter> parameters,
            IReadOnlyDictionary<string, object?>? defaultValues,
            IReadOnlyList<ParameterSettings>? settings,
            out string reportParametersXml,
            out string queryParametersXml)
        {
            var reportParameters =
                new StringBuilder();

            var queryParameters =
                new StringBuilder();

            foreach (var parameter in parameters)
            {
                if (parameter.IsOutput)
                    continue;

                string sqlParameterName =
                    parameter.Name;

                string reportParameterName =
                    sqlParameterName.TrimStart('@');

                // A one-column table-valued parameter is a multi-value parameter of its column's type.
                bool multiValue = parameter.IsTable;

                string rdlType =
                    GetRdlParameterType(
                        multiValue ? parameter.Columns[0].SqlType : parameter.SqlType);

                var setting = settings?.FirstOrDefault(s =>
                    string.Equals(s.SqlName, sqlParameterName, StringComparison.OrdinalIgnoreCase));

                bool required = setting?.IfBlank == BlankValue.Required;
                string prompt = string.IsNullOrWhiteSpace(setting?.DisplayName)
                    ? reportParameterName
                    : setting!.DisplayName.Trim();

                object? defaultValue = null;

                if (defaultValues != null &&
                    defaultValues.TryGetValue(sqlParameterName, out object? given) &&
                    given != null)
                {
                    defaultValue = given;
                }
                else if (setting?.DefaultValue != null && !multiValue)
                {
                    defaultValue = StoredProcedureRunner.ParseValue(parameter.SqlType, setting.DefaultValue);
                }

                string defaultXml = "";

                if (defaultValue != null)
                {
                    defaultXml = $@"
      <DefaultValue>
        <Values>
          <Value>{XmlEscape(FormatRdlValue(defaultValue, rdlType))}</Value>
        </Values>
      </DefaultValue>";
                }
                else if (setting != null && !required && !multiValue)
                {
                    // Starts as NULL (blank), so the report opens with data.
                    defaultXml = @"
      <DefaultValue>
        <Values>
          <Value>=Nothing</Value>
        </Values>
      </DefaultValue>";
                }

                // A dropdown: typed-in values, or the list's dataset (see ListDataSetsXml).
                string validValuesXml = setting?.HasList != true ? ""
                    : setting.ListSource == ListSource.Values
                    ? $@"
      <ValidValues>
        <ParameterValues>{string.Concat(setting.ListRows.Select(row => $@"
          <ParameterValue>
            <Value>{XmlEscape(FormatRdlValue(row.Key!, rdlType))}</Value>
            <Label>{XmlEscape(Convert.ToString(row.Value, System.Globalization.CultureInfo.InvariantCulture) ?? "")}</Label>
          </ParameterValue>"))}
        </ParameterValues>
      </ValidValues>"
                    : $@"
      <ValidValues>
        <DataSetReference>
          <DataSetName>{XmlEscape(ListDataSetName(reportParameterName))}</DataSetName>
          <ValueField>Value</ValueField>
          <LabelField>Label</LabelField>
        </DataSetReference>
      </ValidValues>";

                // Multi-value parameters can't be null: there's at least one value.
                reportParameters.Append($@"
    <ReportParameter Name=""{XmlEscape(reportParameterName)}"">
      <DataType>{rdlType}</DataType>
      <Nullable>{(required || multiValue ? "false" : "true")}</Nullable>{defaultXml}
      <Prompt>{XmlEscape(prompt)}</Prompt>{validValuesXml}{(multiValue ? @"
      <MultiValue>true</MultiValue>" : "")}
    </ReportParameter>");

                string value = multiValue
                    ? $"=Code.FlankTable(Parameters!{reportParameterName}.Value)"
                    : $"=Parameters!{reportParameterName}.Value";

                queryParameters.Append($@"
        <QueryParameter Name=""{XmlEscape(sqlParameterName)}"">
          <Value>{XmlEscape(value)}</Value>
        </QueryParameter>");
            }

            if (reportParameters.Length > 0)
            {
                reportParametersXml = $@"
  <ReportParameters>
{reportParameters}
  </ReportParameters>";
            }
            else
            {
                reportParametersXml = "";
            }

            if (queryParameters.Length > 0)
            {
                queryParametersXml = $@"
      <QueryParameters>
{queryParameters}
      </QueryParameters>";
            }
            else
            {
                queryParametersXml = "";
            }
        }

        private static string BuildReportRdl(
            IReadOnlyList<ReportColumn> columns,
            RdlDataSource dataSource,
            string queryXml,
            string reportParametersXml,
            string extraDataSetsXml = "",
            string codeXml = "")
        {
            TableRdl table =
                BuildTableRdl(columns);

            double reportWidth =
                Math.Max(
                    8.5,
                    columns.Count * 1.5);

            return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<Report
    xmlns=""http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition""
    xmlns:rd=""http://schemas.microsoft.com/SQLServer/reporting/reportdesigner"">

  <DataSources>{dataSource.ToXml()}
  </DataSources>

  <DataSets>
    <DataSet Name=""MainDataSet"">
      <Query>
        <DataSourceName>{XmlEscape(dataSource.Name)}</DataSourceName>
{queryXml}
      </Query>

      <Fields>
{table.Fields}
      </Fields>
    </DataSet>{extraDataSetsXml}
  </DataSets>

{reportParametersXml}
{codeXml}
  <ReportSections>
    <ReportSection>

      <Body>
        <ReportItems>

          <Tablix Name=""MainTable"">

            <TablixBody>

              <TablixColumns>
{table.Columns}
              </TablixColumns>

              <TablixRows>

                <TablixRow>
                  <Height>0.3in</Height>
                  <TablixCells>
{table.HeaderCells}
                  </TablixCells>
                </TablixRow>

                <TablixRow>
                  <Height>0.3in</Height>
                  <TablixCells>
{table.DetailCells}
                  </TablixCells>
                </TablixRow>

              </TablixRows>

            </TablixBody>

            <TablixColumnHierarchy>
              <TablixMembers>
{table.ColumnMembers}
              </TablixMembers>
            </TablixColumnHierarchy>

            <TablixRowHierarchy>
              <TablixMembers>

                <TablixMember>
                  <KeepWithGroup>After</KeepWithGroup>
                  <RepeatOnNewPage>true</RepeatOnNewPage>
                </TablixMember>

                <TablixMember>
                  <Group Name=""Details"" />
                </TablixMember>

              </TablixMembers>
            </TablixRowHierarchy>

            <DataSetName>MainDataSet</DataSetName>

          </Tablix>

        </ReportItems>

        <Height>2in</Height>
        <Style />
      </Body>

      <Width>{reportWidth.ToString(
          System.Globalization.CultureInfo.InvariantCulture)}in</Width>

      <Page>
        <PageHeight>11in</PageHeight>
        <PageWidth>11in</PageWidth>
        <LeftMargin>0.5in</LeftMargin>
        <RightMargin>0.5in</RightMargin>
        <TopMargin>0.5in</TopMargin>
        <BottomMargin>0.5in</BottomMargin>
        <Style />
      </Page>

    </ReportSection>
  </ReportSections>

</Report>";
        }

        private static TableRdl BuildTableRdl(
            IReadOnlyList<ReportColumn> columns)
        {
            var fields = new StringBuilder();
            var tablixColumns = new StringBuilder();
            var headerCells = new StringBuilder();
            var detailCells = new StringBuilder();
            var columnMembers = new StringBuilder();

            int columnNumber = 0;

            // SSRS requires unique field and textbox names, but different
            // columns can map to the same safe name ("a b" and "a_b").
            var usedFieldNames =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var column in columns)
            {
                columnNumber++;

                string sourceName = column.Name;

                if (string.IsNullOrWhiteSpace(sourceName))
                {
                    sourceName =
                        "Column" + columnNumber;
                }

                string fieldName =
                    MakeSafeRdlName(
                        sourceName,
                        columnNumber);

                string baseFieldName = fieldName;

                for (int suffix = 2; !usedFieldNames.Add(fieldName); suffix++)
                {
                    fieldName = baseFieldName + "_" + suffix;
                }

                fields.Append($@"
        <Field Name=""{XmlEscape(fieldName)}"">
          <DataField>{XmlEscape(sourceName)}</DataField>
        </Field>");

                tablixColumns.Append(@"
                <TablixColumn>
                  <Width>1.5in</Width>
                </TablixColumn>");

                headerCells.Append($@"
                <TablixCell>
                  <CellContents>
                    <Textbox Name=""Header_{XmlEscape(fieldName)}"">
                      <CanGrow>true</CanGrow>
                      <KeepTogether>true</KeepTogether>
                      <Paragraphs>
                        <Paragraph>
                          <TextRuns>
                            <TextRun>
                              <Value>{XmlEscape(sourceName)}</Value>
                              <Style>
                                <FontWeight>Bold</FontWeight>
                              </Style>
                            </TextRun>
                          </TextRuns>
                          <Style />
                        </Paragraph>
                      </Paragraphs>
                      <Style />
                    </Textbox>
                  </CellContents>
                </TablixCell>");

                detailCells.Append($@"
                <TablixCell>
                  <CellContents>
                    <Textbox Name=""Value_{XmlEscape(fieldName)}"">
                      <CanGrow>true</CanGrow>
                      <KeepTogether>true</KeepTogether>
                      <Paragraphs>
                        <Paragraph>
                          <TextRuns>
                            <TextRun>
                              <Value>=Fields!{XmlEscape(fieldName)}.Value</Value>
                              <Style />
                            </TextRun>
                          </TextRuns>
                          <Style />
                        </Paragraph>
                      </Paragraphs>
                      <Style />
                    </Textbox>
                  </CellContents>
                </TablixCell>");

                columnMembers.Append(@"
                <TablixMember />");
            }

            return new TableRdl
            {
                Fields = fields.ToString(),
                Columns = tablixColumns.ToString(),
                HeaderCells = headerCells.ToString(),
                DetailCells = detailCells.ToString(),
                ColumnMembers = columnMembers.ToString()
            };
        }

        private void DeployReport(
            string rdl,
            string reportFolder,
            string reportName)
        {
            var rs = CreateSsrsConnection();

            byte[] definition =
                Encoding.UTF8.GetBytes(rdl);

            Warning[] warnings;

            try
            {
                rs.CreateCatalogItem(
                    "Report",
                    reportName,
                    reportFolder,
                    true,
                    definition,
                    null,
                    out warnings);
            }
            catch (System.Web.Services.Protocols.SoapException ex)
            {
                throw new SsrsException(
                    $"Could not create report '{reportName}' " +
                    $"in SSRS folder '{reportFolder}'. " +
                    GetSoapErrorMessage(ex),
                    ex);
            }
            catch (Exception ex)
            {
                throw new SsrsException(
                    $"Could not create report '{reportName}' " +
                    $"in SSRS folder '{reportFolder}'.",
                    ex);
            }

            if (warnings != null)
            {
                foreach (var warning in warnings)
                {
                    Console.WriteLine(
                        "WARNING: " + warning.Message);
                }
            }
        }

        private ReportService2010.ReportingService2010 CreateSsrsConnection()
        {
            return new ReportService2010.ReportingService2010
            {
                Url = _reportServerUrl + "/ReportService2010.asmx",
                Credentials = _credentials ?? CredentialCache.DefaultCredentials
            };
        }

        private string GetReportUrl(
            string reportFolder,
            string reportName)
        {
            string reportPath =
                reportFolder.TrimEnd('/') +
                "/" +
                reportName;

            string encodedPath = string.Join(
                "",
                reportPath
                    .Split('/')
                    .Where(x =>
                        !string.IsNullOrEmpty(x))
                    .Select(x =>
                        "/" + Uri.EscapeDataString(x)));

            return _reportPortalUrl +
                "/report" +
                encodedPath;
        }

        private static string GetSoapErrorMessage(
            System.Web.Services.Protocols.SoapException ex)
        {
            if (ex.Detail != null &&
                !string.IsNullOrWhiteSpace(
                    ex.Detail.InnerText))
            {
                return ex.Detail.InnerText;
            }

            return ex.Message;
        }

        private static string XmlEscape(
            string value)
        {
            return System.Security
                .SecurityElement
                .Escape(value);
        }

        // RDL parameter values are invariant text: 2024-01-02T00:00:00, true, 1.5.
        internal static string FormatRdlValue(object value, string rdlType)
        {
            switch (value)
            {
                case DateTime date:
                    return date.ToString("yyyy-MM-ddTHH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                case bool flag:
                    return flag ? "true" : "false";
                case IFormattable formattable:
                    return formattable.ToString(null, System.Globalization.CultureInfo.InvariantCulture);
                default:
                    return value.ToString() ?? "";
            }
        }

        internal static string GetRdlParameterType(
            string sqlType)
        {
            switch (sqlType.ToLowerInvariant())
            {
                case "tinyint":
                case "smallint":
                case "int":
                    return "Integer";

                case "bigint":
                    // SSRS Integer is Int32.
                    return "String";

                case "decimal":
                case "numeric":
                case "money":
                case "smallmoney":
                case "float":
                case "real":
                    return "Float";

                case "bit":
                    return "Boolean";

                case "date":
                case "datetime":
                case "datetime2":
                case "smalldatetime":
                    return "DateTime";

                case "char":
                case "varchar":
                case "nchar":
                case "nvarchar":
                case "text":
                case "ntext":
                case "uniqueidentifier":
                case "time":
                case "datetimeoffset":
                    return "String";

                default:
                    throw new NotSupportedException(
                        $"SQL parameter type '{sqlType}' " +
                        "is not supported.");
            }
        }

        internal static string MakeSafeRdlName(
            string name,
            int fallbackNumber)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "Column" +
                    fallbackNumber;
            }

            var sb = new StringBuilder();

            foreach (char c in name)
            {
                if (char.IsLetterOrDigit(c) ||
                    c == '_')
                {
                    sb.Append(c);
                }
                else
                {
                    sb.Append('_');
                }
            }

            if (sb.Length == 0)
            {
                return "Column" +
                    fallbackNumber;
            }

            if (char.IsDigit(sb[0]))
            {
                sb.Insert(0, '_');
            }

            return sb.ToString();
        }

        public List<string> GetFolders()
        {
            try
            {
                var rs =
                    CreateSsrsConnection();

                var items =
                    rs.ListChildren("/", true);

                var folders = items
                    .Where(x =>
                        x.TypeName == "Folder")
                    .Select(x => x.Path)
                    .ToList();

                if (!folders.Contains("/"))
                {
                    folders.Insert(0, "/");
                }
                return folders;
            }
            catch (Exception ex)
            {
                throw new SsrsException(
                    "Could not retrieve folders " +
                    "from the SSRS server.",
                    ex);
            }
        }

        public List<string> GetSharedDataSources()
        {
            try
            {
                var rs =
                    CreateSsrsConnection();

                var items =
                    rs.ListChildren("/", true);

                return items
                    .Where(x =>
                        x.TypeName == "DataSource")
                    .Select(x => x.Path)
                    .ToList();
            }
            catch (Exception ex)
            {
                throw new SsrsException(
                    "Could not retrieve shared data sources " +
                    "from the SSRS server.",
                    ex);
            }
        }
    }
}