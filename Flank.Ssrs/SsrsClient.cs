using Flank.Ssrs.ReportService2010;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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

        private class SprocParameter
        {
            public string Name { get; set; }
            public string SqlType { get; set; }
            public short MaxLength { get; set; }
            public byte Precision { get; set; }
            public byte Scale { get; set; }
            public bool IsOutput { get; set; }
        }

        private class ReportColumn
        {
            public string Name { get; set; }
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
        public string CreateReport(
            string sql,
            SqlConnection connection,
            string sharedDataSourcePath,
            string reportFolder,
            string reportName)
        {
            var columns = DiscoverQueryColumns(
                sql,
                connection);

            string queryXml = $@"
        <CommandText>{XmlEscape(sql)}</CommandText>";

            string rdl = BuildReportRdl(
                columns,
                sharedDataSourcePath,
                queryXml,
                null);

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
            SqlConnection connection,
            string sharedDataSourcePath,
            string reportFolder,
            string reportName)
        {
            int objectId = GetStoredProcedureObjectId(
                procedureName,
                connection);

            var parameters = DiscoverStoredProcedureParameters(
                objectId,
                connection);

            var columns = DiscoverStoredProcedureColumns(
                objectId,
                procedureName,
                connection);

            string reportParametersXml;
            string queryParametersXml;

            BuildStoredProcedureParameters(
                parameters,
                out reportParametersXml,
                out queryParametersXml);

            string queryXml = $@"
        <CommandType>StoredProcedure</CommandType>
        <CommandText>{XmlEscape(procedureName)}</CommandText>
{queryParametersXml}";

            string rdl = BuildReportRdl(
                columns,
                sharedDataSourcePath,
                queryXml,
                reportParametersXml);

            DeployReport(
                rdl,
                reportFolder,
                reportName);

            return GetReportUrl(
                reportFolder,
                reportName);
        }

        private List<ReportColumn> DiscoverQueryColumns(
            string sql,
            SqlConnection connection)
        {
            var columns = new List<ReportColumn>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = sql;

                using (var reader = command.ExecuteReader(
                    System.Data.CommandBehavior.SchemaOnly))
                {
                    var schema = reader.GetSchemaTable();

                    foreach (System.Data.DataRow row in schema.Rows)
                    {
                        columns.Add(new ReportColumn
                        {
                            Name = (string)row["ColumnName"]
                        });
                    }
                }
            }

            if (columns.Count == 0)
            {
                throw new InvalidOperationException(
                    "The query does not return any columns.");
            }

            return columns;
        }

        private int GetStoredProcedureObjectId(
            string procedureName,
            SqlConnection connection)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT OBJECT_ID(@ProcedureName, 'P')";

                command.Parameters.AddWithValue(
                    "@ProcedureName",
                    procedureName);

                var result = command.ExecuteScalar();

                if (result == null ||
                    result == DBNull.Value)
                {
                    throw new InvalidOperationException(
                        $"Stored procedure '{procedureName}' not found.");
                }

                return (int)result;
            }
        }

        private List<SprocParameter> DiscoverStoredProcedureParameters(
            int objectId,
            SqlConnection connection)
        {
            var parameters = new List<SprocParameter>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT
    p.name,
    t.name AS type_name,
    p.max_length,
    p.precision,
    p.scale,
    p.is_output
FROM sys.parameters p
JOIN sys.types t
    ON p.user_type_id = t.user_type_id
WHERE p.object_id = @ObjectId
ORDER BY p.parameter_id;";

                command.Parameters.AddWithValue(
                    "@ObjectId",
                    objectId);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        parameters.Add(new SprocParameter
                        {
                            Name = reader.GetString(0),
                            SqlType = reader.GetString(1),
                            MaxLength = reader.GetInt16(2),
                            Precision = reader.GetByte(3),
                            Scale = reader.GetByte(4),
                            IsOutput = reader.GetBoolean(5)
                        });
                    }
                }
            }

            return parameters;
        }

        private List<ReportColumn> DiscoverStoredProcedureColumns(
            int objectId,
            string procedureName,
            SqlConnection connection)
        {
            var columns = new List<ReportColumn>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT
    name,
    system_type_name,
    is_nullable,
    column_ordinal
FROM sys.dm_exec_describe_first_result_set_for_object(
    @ObjectId,
    NULL
)
WHERE is_hidden = 0
ORDER BY column_ordinal;";

                command.Parameters.AddWithValue(
                    "@ObjectId",
                    objectId);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columns.Add(new ReportColumn
                        {
                            Name = reader.IsDBNull(0)
                                ? null
                                : reader.GetString(0)
                        });
                    }
                }
            }

            if (columns.Count == 0)
            {
                throw new InvalidOperationException(
                    $"Stored procedure '{procedureName}' " +
                    "does not return a result set.");
            }

            return columns;
        }

        private static void BuildStoredProcedureParameters(
            List<SprocParameter> parameters,
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

                string rdlType =
                    GetRdlParameterType(
                        parameter.SqlType);

                reportParameters.Append($@"
    <ReportParameter Name=""{XmlEscape(reportParameterName)}"">
      <DataType>{rdlType}</DataType>
      <Nullable>true</Nullable>
      <Prompt>{XmlEscape(reportParameterName)}</Prompt>
    </ReportParameter>");

                queryParameters.Append($@"
        <QueryParameter Name=""{XmlEscape(sqlParameterName)}"">
          <Value>=Parameters!{XmlEscape(reportParameterName)}.Value</Value>
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
            List<ReportColumn> columns,
            string sharedDataSourcePath,
            string queryXml,
            string reportParametersXml)
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

  <DataSources>
    <DataSource Name=""SharedDataSource"">
      <DataSourceReference>{XmlEscape(sharedDataSourcePath)}</DataSourceReference>
    </DataSource>
  </DataSources>

  <DataSets>
    <DataSet Name=""MainDataSet"">
      <Query>
        <DataSourceName>SharedDataSource</DataSourceName>
{queryXml}
      </Query>

      <Fields>
{table.Fields}
      </Fields>
    </DataSet>
  </DataSets>

{reportParametersXml}

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
            List<ReportColumn> columns)
        {
            var fields = new StringBuilder();
            var tablixColumns = new StringBuilder();
            var headerCells = new StringBuilder();
            var detailCells = new StringBuilder();
            var columnMembers = new StringBuilder();

            int columnNumber = 0;

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

        private static string GetRdlParameterType(
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

        private static string MakeSafeRdlName(
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

                return items
                    .Where(x =>
                        x.TypeName == "Folder")
                    .Select(x => x.Path)
                    .ToList();
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