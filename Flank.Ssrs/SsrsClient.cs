using Flank.Ssrs.ReportService2010;
using System;
using System.Data.SqlClient;
using System.Net;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace Flank.Ssrs
{
    public class SsrsClient
    {
        private readonly string _reportServerUrl;
        private readonly string _reportPortalUrl;

        public SsrsClient(
            string reportServerUrl,
            string reportPortalUrl)
        {
            _reportServerUrl = reportServerUrl.TrimEnd('/');
            _reportPortalUrl = reportPortalUrl.TrimEnd('/');
        }

        private ReportService2010.ReportingService2010 CreateSsrsConnection()
        {
            return new ReportService2010.ReportingService2010
            {
                Url = _reportServerUrl + "/ReportService2010.asmx",
                Credentials = System.Net.CredentialCache.DefaultCredentials
            };
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

        private class SprocColumn
        {
            public string Name { get; set; }
            public string SqlType { get; set; }
            public bool? IsNullable { get; set; }
            public int Ordinal { get; set; }
        }
        private string GetReportUrl(
        string reportFolder,
        string reportName)
        {
            string reportPath =
                reportFolder.TrimEnd('/') + "/" + reportName;

            string encodedPath = string.Join(
                "",
                reportPath
                    .Split('/')
                    .Where(x => !string.IsNullOrEmpty(x))
                    .Select(x => "/" + Uri.EscapeDataString(x)));

            return _reportPortalUrl + "/report" + encodedPath;
        }
        public void TestConnection()
        {
            var rs = CreateSsrsConnection();

            rs.Url =
                "http://localhost/ReportServer/ReportService2010.asmx";

            rs.Credentials = CredentialCache.DefaultCredentials;

            var items = rs.ListChildren("/", false);

            Console.WriteLine($"Found {items.Length} items.");

            foreach (var item in items)
            {
                Console.WriteLine($"{item.TypeName}: {item.Name}");
            }
        }

        public string CreateReport(
            string sql,
            System.Data.SqlClient.SqlConnection connection,
            string sharedDataSourcePath,
            string reportFolder,
            string reportName)
        {
            // 1. Discover the output columns without executing the query normally.
            var columns = new List<string>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = sql;

                using (var reader = command.ExecuteReader(
                    System.Data.CommandBehavior.SchemaOnly))
                {
                    var schema = reader.GetSchemaTable();

                    foreach (System.Data.DataRow row in schema.Rows)
                    {
                        columns.Add((string)row["ColumnName"]);
                    }
                }
            }

            if (columns.Count == 0)
                throw new InvalidOperationException(
                    "The query does not return any columns.");

            // 2. Build the dynamic pieces of the RDL.
            var fieldXml = new StringBuilder();
            var columnXml = new StringBuilder();
            var headerCellXml = new StringBuilder();
            var detailCellXml = new StringBuilder();
            var columnMemberXml = new StringBuilder();

            foreach (var column in columns)
            {
                string name = XmlEscape(column);

                fieldXml.Append($@"
        <Field Name=""{name}"">
          <DataField>{name}</DataField>
        </Field>");

                columnXml.Append(@"
                <TablixColumn>
                  <Width>1.5in</Width>
                </TablixColumn>");

                headerCellXml.Append($@"
                <TablixCell>
                  <CellContents>
                    <Textbox Name=""Header_{name}"">
                      <Paragraphs>
                        <Paragraph>
                          <TextRuns>
                            <TextRun>
                              <Value>{name}</Value>
                              <Style>
                                <FontWeight>Bold</FontWeight>
                              </Style>
                            </TextRun>
                          </TextRuns>
                        </Paragraph>
                      </Paragraphs>
                    </Textbox>
                  </CellContents>
                </TablixCell>");

                detailCellXml.Append($@"
                <TablixCell>
                  <CellContents>
                    <Textbox Name=""Value_{name}"">
                      <Paragraphs>
                        <Paragraph>
                          <TextRuns>
                            <TextRun>
                              <Value>=Fields!{name}.Value</Value>
                            </TextRun>
                          </TextRuns>
                        </Paragraph>
                      </Paragraphs>
                    </Textbox>
                  </CellContents>
                </TablixCell>");

                columnMemberXml.Append("<TablixMember />");
            }

            // 3. Generate the RDL.
            string rdl = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<Report xmlns=""http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition"">

  <DataSources>
    <DataSource Name=""SharedDataSource"">
      <DataSourceReference>{XmlEscape(sharedDataSourcePath)}</DataSourceReference>
    </DataSource>
  </DataSources>

  <DataSets>
    <DataSet Name=""MainDataSet"">
      <Query>
        <DataSourceName>SharedDataSource</DataSourceName>
        <CommandText>{XmlEscape(sql)}</CommandText>
      </Query>

      <Fields>
        {fieldXml}
      </Fields>
    </DataSet>
  </DataSets>

  <ReportSections>
    <ReportSection>
      <Body>
        <ReportItems>

          <Tablix Name=""MainTable"">
            <TablixBody>

              <TablixColumns>
                {columnXml}
              </TablixColumns>

              <TablixRows>

                <!-- Header -->
                <TablixRow>
                  <Height>0.3in</Height>
                  <TablixCells>
                    {headerCellXml}
                  </TablixCells>
                </TablixRow>

                <!-- Detail -->
                <TablixRow>
                  <Height>0.3in</Height>
                  <TablixCells>
                    {detailCellXml}
                  </TablixCells>
                </TablixRow>

              </TablixRows>
            </TablixBody>

            <TablixColumnHierarchy>
              <TablixMembers>
                {columnMemberXml}
              </TablixMembers>
            </TablixColumnHierarchy>

            <TablixRowHierarchy>
              <TablixMembers>
                <TablixMember />
                <TablixMember>
                  <Group Name=""Details"" />
                </TablixMember>
              </TablixMembers>
            </TablixRowHierarchy>

            <DataSetName>MainDataSet</DataSetName>
          </Tablix>

        </ReportItems>

        <Height>2in</Height>
      </Body>

      <Width>{Math.Max(8.5, columns.Count * 1.5)}in</Width>

      <Page>
        <PageHeight>11in</PageHeight>
        <PageWidth>11in</PageWidth>
        <LeftMargin>0.5in</LeftMargin>
        <RightMargin>0.5in</RightMargin>
        <TopMargin>0.5in</TopMargin>
        <BottomMargin>0.5in</BottomMargin>
      </Page>

    </ReportSection>
  </ReportSections>

</Report>";

            // 4. Connect to SSRS.
            var rs = CreateSsrsConnection();

            // 5. Deploy.
            byte[] definition = Encoding.UTF8.GetBytes(rdl);

            Warning[] warnings;

            rs.CreateCatalogItem(
                "Report",
                reportName,
                reportFolder,
                true,
                definition,
                null,
                out warnings);

            Console.WriteLine(
                $"Created {reportName} with {columns.Count} columns.");

            if (warnings != null)
            {
                foreach (var warning in warnings)
                    Console.WriteLine("WARNING: " + warning.Message);
            }
            return GetReportUrl(reportFolder, reportName);
        }

        private static string XmlEscape(string value)
        {
            return System.Security.SecurityElement.Escape(value);
        }
        public string CreateReportFromStoredProcedure(
             string procedureName,
             SqlConnection connection,
             string sharedDataSourcePath,
             string reportFolder,
             string reportName)
        {
            // ------------------------------------------------------------
            // 1. Resolve the stored procedure
            // ------------------------------------------------------------

            int objectId;

            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
            SELECT OBJECT_ID(@ProcedureName, 'P')";

                command.Parameters.AddWithValue("@ProcedureName", procedureName);

                var result = command.ExecuteScalar();

                if (result == null || result == DBNull.Value)
                    throw new InvalidOperationException(
                        $"Stored procedure '{procedureName}' not found.");

                objectId = (int)result;
            }


            // ------------------------------------------------------------
            // 2. Discover stored procedure parameters
            // ------------------------------------------------------------

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

                command.Parameters.AddWithValue("@ObjectId", objectId);

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


            // ------------------------------------------------------------
            // 3. Discover first result-set columns
            // ------------------------------------------------------------

            var columns = new List<SprocColumn>();

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

                command.Parameters.AddWithValue("@ObjectId", objectId);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columns.Add(new SprocColumn
                        {
                            Name = reader.IsDBNull(0)
                                ? null
                                : reader.GetString(0),

                            SqlType = reader.IsDBNull(1)
                                ? null
                                : reader.GetString(1),

                            IsNullable = reader.IsDBNull(2)
                                ? (bool?)null
                                : reader.GetBoolean(2),

                            Ordinal = reader.GetInt32(3)
                        });
                    }
                }
            }

            if (columns.Count == 0)
                throw new InvalidOperationException(
                    $"Stored procedure '{procedureName}' does not return a result set.");


            // ------------------------------------------------------------
            // 4. Generate RDL parameter XML
            // ------------------------------------------------------------

            var reportParameterXml = new StringBuilder();
            var dataSetParameterXml = new StringBuilder();

            foreach (var parameter in parameters)
            {
                // Ignore SQL output parameters for the report-input UI for now.
                if (parameter.IsOutput)
                    continue;

                string sqlParameterName = parameter.Name;

                // SSRS report parameter names don't include @.
                string reportParameterName =
                    sqlParameterName.TrimStart('@');

                string rdlType = GetRdlParameterType(parameter.SqlType);

                reportParameterXml.Append($@"
    <ReportParameter Name=""{XmlEscape(reportParameterName)}"">
      <DataType>{rdlType}</DataType>
      <Prompt>{XmlEscape(reportParameterName)}</Prompt>
    </ReportParameter>");

                dataSetParameterXml.Append($@"
        <QueryParameter Name=""{XmlEscape(sqlParameterName)}"">
          <Value>=Parameters!{XmlEscape(reportParameterName)}.Value</Value>
        </QueryParameter>");
            }


            // ------------------------------------------------------------
            // 5. Generate dataset fields + tablix
            // ------------------------------------------------------------

            var fieldXml = new StringBuilder();
            var tablixColumnXml = new StringBuilder();
            var headerCellXml = new StringBuilder();
            var detailCellXml = new StringBuilder();
            var columnMemberXml = new StringBuilder();

            int columnNumber = 0;

            foreach (var column in columns)
            {
                columnNumber++;

                // SQL Server can technically return unnamed expressions.
                // Give them a usable RDL field name.
                string sourceName = column.Name;

                if (string.IsNullOrWhiteSpace(sourceName))
                    sourceName = "Column" + columnNumber;

                // Use a safe internal identifier for RDL object/field names.
                string fieldName = MakeSafeRdlName(sourceName, columnNumber);

                fieldXml.Append($@"
        <Field Name=""{XmlEscape(fieldName)}"">
          <DataField>{XmlEscape(sourceName)}</DataField>
        </Field>");

                tablixColumnXml.Append(@"
                <TablixColumn>
                  <Width>1.5in</Width>
                </TablixColumn>");

                headerCellXml.Append($@"
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

                detailCellXml.Append($@"
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

                columnMemberXml.Append(@"
                <TablixMember />");
            }


            // ------------------------------------------------------------
            // 6. Generate complete RDL
            // ------------------------------------------------------------

            string reportParametersSection = "";

            if (reportParameterXml.Length > 0)
            {
                reportParametersSection = $@"
  <ReportParameters>
{reportParameterXml}
  </ReportParameters>";
            }

            string queryParametersSection = "";

            if (dataSetParameterXml.Length > 0)
            {
                queryParametersSection = $@"
      <QueryParameters>
{dataSetParameterXml}
      </QueryParameters>";
            }

            double reportWidth = Math.Max(
                8.5,
                columns.Count * 1.5);

            string rdl = $@"<?xml version=""1.0"" encoding=""utf-8""?>
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
        <CommandType>StoredProcedure</CommandType>
        <CommandText>{XmlEscape(procedureName)}</CommandText>
{queryParametersSection}
      </Query>

      <Fields>
{fieldXml}
      </Fields>
    </DataSet>
  </DataSets>

{reportParametersSection}

  <ReportSections>
    <ReportSection>

      <Body>
        <ReportItems>

          <Tablix Name=""MainTable"">

            <TablixBody>

              <TablixColumns>
{tablixColumnXml}
              </TablixColumns>

              <TablixRows>

                <TablixRow>
                  <Height>0.3in</Height>
                  <TablixCells>
{headerCellXml}
                  </TablixCells>
                </TablixRow>

                <TablixRow>
                  <Height>0.3in</Height>
                  <TablixCells>
{detailCellXml}
                  </TablixCells>
                </TablixRow>

              </TablixRows>

            </TablixBody>

            <TablixColumnHierarchy>
              <TablixMembers>
{columnMemberXml}
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


            // ------------------------------------------------------------
            // 7. Connect to SSRS
            // ------------------------------------------------------------

            var rs = CreateSsrsConnection();


            // ------------------------------------------------------------
            // 8. Deploy report
            // ------------------------------------------------------------

            byte[] definition = Encoding.UTF8.GetBytes(rdl);

            Warning[] warnings;

            rs.CreateCatalogItem(
                "Report",
                reportName,
                reportFolder,
                true,
                definition,
                null,
                out warnings);

            Console.WriteLine(
                $"Created '{reportName}' from '{procedureName}'.");

            Console.WriteLine(
                $"Inputs: {parameters.Count}, Outputs: {columns.Count}");

            if (warnings != null)
            {
                foreach (var warning in warnings)
                    Console.WriteLine("WARNING: " + warning.Message);
            }
            return GetReportUrl(reportFolder, reportName);
        }
        private static string GetRdlParameterType(string sqlType)
        {
            switch (sqlType.ToLowerInvariant())
            {
                case "tinyint":
                case "smallint":
                case "int":
                    return "Integer";

                case "bigint":
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

                default:
                    return "String";
            }
        }

        private static string MakeSafeRdlName(
            string name,
            int fallbackNumber)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Column" + fallbackNumber;

            var sb = new StringBuilder();

            foreach (char c in name)
            {
                if (char.IsLetterOrDigit(c) || c == '_')
                    sb.Append(c);
                else
                    sb.Append('_');
            }

            if (sb.Length == 0)
                return "Column" + fallbackNumber;

            if (char.IsDigit(sb[0]))
                sb.Insert(0, '_');

            return sb.ToString();
        }

        public List<string> GetFolders()
        {
            var rs = CreateSsrsConnection();

            var items = rs.ListChildren("/", true);

            return items
                .Where(x => x.TypeName == "Folder")
                .Select(x => x.Path)
                .ToList();
        }

        public List<string> GetSharedDataSources()
        {
            var rs = CreateSsrsConnection();

            var items = rs.ListChildren("/", true);

            return items
                .Where(x => x.TypeName == "DataSource")
                .Select(x => x.Path)
                .ToList();
        }
    }
}