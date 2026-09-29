using Flank.Ssrs.ReportService2010;
using System;
using System.Net;

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
        public void CreateTestReport()
        {
            var rs = new Flank.Ssrs.ReportService2010.ReportingService2010();

            rs.Url = "http://localhost/ReportServer/ReportService2010.asmx";
            rs.Credentials = System.Net.CredentialCache.DefaultCredentials;

            string rdl = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Report xmlns=""http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition""
        xmlns:rd=""http://schemas.microsoft.com/SQLServer/reporting/reportdesigner"">
  <ReportSections>
    <ReportSection>
      <Body>
        <ReportItems>
          <Textbox Name=""HelloTextbox"">
            <CanGrow>true</CanGrow>
            <KeepTogether>true</KeepTogether>
            <Paragraphs>
              <Paragraph>
                <TextRuns>
                  <TextRun>
                    <Value>Hello from Flank</Value>
                    <Style>
                      <FontSize>20pt</FontSize>
                    </Style>
                  </TextRun>
                </TextRuns>
                <Style />
              </Paragraph>
            </Paragraphs>
            <Top>0.5in</Top>
            <Left>0.5in</Left>
            <Height>0.4in</Height>
            <Width>3in</Width>
            <Style />
          </Textbox>
        </ReportItems>
        <Height>2in</Height>
        <Style />
      </Body>
      <Width>8.5in</Width>
      <Page>
        <PageHeight>11in</PageHeight>
        <PageWidth>8.5in</PageWidth>
        <LeftMargin>1in</LeftMargin>
        <RightMargin>1in</RightMargin>
        <TopMargin>1in</TopMargin>
        <BottomMargin>1in</BottomMargin>
        <Style />
      </Page>
    </ReportSection>
  </ReportSections>
</Report>";

            byte[] definition = System.Text.Encoding.UTF8.GetBytes(rdl);

            Warning[] warnings;

            rs.CreateCatalogItem(
                "Report",
                "Flank Test Report",
                "/",
                true,
                definition,
                null,
                out warnings);

            Console.WriteLine("Report created.");

            if (warnings != null)
            {
                foreach (var warning in warnings)
                    Console.WriteLine(warning.Message);
            }
        }
        public void CreateBikesReport()
        {
            var rs = new Flank.Ssrs.ReportService2010.ReportingService2010
            {
                Url = "http://localhost/ReportServer/ReportService2010.asmx",
                Credentials = System.Net.CredentialCache.DefaultCredentials
            };

            string rdl = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Report xmlns=""http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition"">

  <DataSources>
    <DataSource Name=""HardcodedTest"">
      <DataSourceReference>/HardcodedTest</DataSourceReference>
    </DataSource>
  </DataSources>

  <DataSets>
    <DataSet Name=""Bikes"">
      <Query>
        <DataSourceName>HardcodedTest</DataSourceName>
        <CommandText>select * from bikes;</CommandText>
      </Query>
      <Fields>
        <Field Name=""bike_id""><DataField>bike_id</DataField></Field>
        <Field Name=""battery_percentage""><DataField>battery_percentage</DataField></Field>
        <Field Name=""duration_checked_out""><DataField>duration_checked_out</DataField></Field>
        <Field Name=""near_home_or_charge""><DataField>near_home_or_charge</DataField></Field>
        <Field Name=""current_user_id""><DataField>current_user_id</DataField></Field>
        <Field Name=""needs_pickup""><DataField>needs_pickup</DataField></Field>
      </Fields>
    </DataSet>
  </DataSets>

  <ReportSections>
    <ReportSection>
      <Body>
        <ReportItems>

          <Tablix Name=""BikesTable"">
            <TablixBody>
              <TablixColumns>
                <TablixColumn><Width>1.2in</Width></TablixColumn>
                <TablixColumn><Width>1.5in</Width></TablixColumn>
                <TablixColumn><Width>1.5in</Width></TablixColumn>
                <TablixColumn><Width>1.5in</Width></TablixColumn>
                <TablixColumn><Width>1.5in</Width></TablixColumn>
                <TablixColumn><Width>1.2in</Width></TablixColumn>
              </TablixColumns>

              <TablixRows>
                <TablixRow>
                  <Height>0.3in</Height>
                  <TablixCells>
                    <TablixCell><CellContents><Textbox Name=""bike_id""><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!bike_id.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs></Textbox></CellContents></TablixCell>
                    <TablixCell><CellContents><Textbox Name=""battery_percentage""><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!battery_percentage.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs></Textbox></CellContents></TablixCell>
                    <TablixCell><CellContents><Textbox Name=""duration_checked_out""><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!duration_checked_out.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs></Textbox></CellContents></TablixCell>
                    <TablixCell><CellContents><Textbox Name=""near_home_or_charge""><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!near_home_or_charge.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs></Textbox></CellContents></TablixCell>
                    <TablixCell><CellContents><Textbox Name=""current_user_id""><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!current_user_id.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs></Textbox></CellContents></TablixCell>
                    <TablixCell><CellContents><Textbox Name=""needs_pickup""><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!needs_pickup.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs></Textbox></CellContents></TablixCell>
                  </TablixCells>
                </TablixRow>
              </TablixRows>
            </TablixBody>

            <TablixColumnHierarchy>
              <TablixMembers>
                <TablixMember/><TablixMember/><TablixMember/>
                <TablixMember/><TablixMember/><TablixMember/>
              </TablixMembers>
            </TablixColumnHierarchy>

            <TablixRowHierarchy>
              <TablixMembers>
                <TablixMember>
                  <Group Name=""Details"" />
                </TablixMember>
              </TablixMembers>
            </TablixRowHierarchy>

            <DataSetName>Bikes</DataSetName>
          </Tablix>

        </ReportItems>
        <Height>2in</Height>
      </Body>

      <Width>9in</Width>

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

            byte[] definition = System.Text.Encoding.UTF8.GetBytes(rdl);

            Warning[] warnings;

            rs.CreateCatalogItem(
                "Report",
                "Bikes",
                "/",
                true,
                definition,
                null,
                out warnings);

            Console.WriteLine("Bikes report created.");

            if (warnings != null)
                foreach (var warning in warnings)
                    Console.WriteLine("WARNING: " + warning.Message);
        }
    }
}