using Flank.Excel;

ExcelExporter.Generate(
    @"..\..\..\..\Flank.Excel\template.xlsx",
    @"C:\temp\flank-test.xlsx");

Console.WriteLine("Done!");