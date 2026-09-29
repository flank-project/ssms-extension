using Flank.Ssrs.ReportService2010;
using System;
using System.Data.SqlClient;
using System.Net;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Flank.Ssrs
{
    public class SsrsClient
    {
        public void TestConnection()
        {
            var rs = new Flank.Ssrs.ReportService2010.ReportingService2010();

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

        public void CreateReport(
            string sql,
            System.Data.SqlClient.SqlConnection connection,
            string sharedDataSourcePath,
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
            var rs = new Flank.Ssrs.ReportService2010.ReportingService2010
            {
                Url = "http://localhost/ReportServer/ReportService2010.asmx",
                Credentials = System.Net.CredentialCache.DefaultCredentials
            };

            // 5. Deploy.
            byte[] definition = Encoding.UTF8.GetBytes(rdl);

            Warning[] warnings;

            rs.CreateCatalogItem(
                "Report",
                reportName,
                "/",
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
        }

        private static string XmlEscape(string value)
        {
            return System.Security.SecurityElement.Escape(value);
        }
    }
}