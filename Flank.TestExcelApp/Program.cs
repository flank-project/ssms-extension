using Flank.Excel;

ExcelExporter.Generate(
    @"..\..\..\..\Flank.Excel\template.xlsx",
    @"C:\temp\flank-test.xlsx", "select * from leagues");

Console.WriteLine("Done!");