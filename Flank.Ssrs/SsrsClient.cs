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
    }
}