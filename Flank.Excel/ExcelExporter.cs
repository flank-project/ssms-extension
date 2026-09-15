namespace Flank.Excel;

public static class ExcelExporter
{
    public static void Generate(string templatePath, string outputPath)
    {
        File.Copy(templatePath, outputPath, overwrite: true);
    }
}