using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Flank.Excel;

public static class ExcelExporter
{
    public static void Generate(string templatePath, string outputPath)
    {
        File.Copy(templatePath, outputPath, overwrite: true);

        using var workbook = ZipFile.Open(outputPath, ZipArchiveMode.Read);

        // Find customXml/item1.xml inside the XLSX
        var item = workbook.GetEntry("customXml/item1.xml")
            ?? throw new Exception("Could not find customXml/item1.xml");

        XDocument xml;

        using (var stream = item.Open())
        {
            xml = XDocument.Load(stream);
        }

        // <DataMashup> contains the Base64-encoded Power Query package
        var dataMashup = xml.Root
            ?? throw new Exception("Could not find DataMashup");

        byte[] mashupBytes = Convert.FromBase64String(dataMashup.Value);

        // Bytes 4-7 contain the length of the embedded ZIP
        int zipLength = BitConverter.ToInt32(mashupBytes, 4);

        Console.WriteLine($"Embedded ZIP length: {zipLength}");

        // ZIP begins at byte 8
        byte[] zipBytes = mashupBytes
            .Skip(8)
            .Take(zipLength)
            .ToArray();

        // Open the embedded ZIP in memory
        using var memory = new MemoryStream(zipBytes);
        using var mashupZip = new ZipArchive(memory, ZipArchiveMode.Read);

        var section = mashupZip.GetEntry("Formulas/Section1.m")
            ?? throw new Exception("Could not find Formulas/Section1.m");

        using var reader = new StreamReader(section.Open(), Encoding.UTF8);

        string mCode = reader.ReadToEnd();

        Console.WriteLine("----- Section1.m -----");
        Console.WriteLine(mCode);
    }
}