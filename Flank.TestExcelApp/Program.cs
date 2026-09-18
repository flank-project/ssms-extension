using Flank.Excel;

ExcelExporter.Generate(
    @"C:\temp\flank-test.xlsx",
    "select * from leagues",
    "fomf-sandbox-ssdb.database.windows.net",
    "fomf-ss-sandbox-db");

Console.WriteLine("Done!");