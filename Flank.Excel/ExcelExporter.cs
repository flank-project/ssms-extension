using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Flank.Excel;

public static class ExcelExporter
{
    public static void Generate(
        string templatePath,
        string outputPath,
        string sql)
    {
        File.Copy(templatePath, outputPath, overwrite: true);

        using var workbook = ZipFile.Open(
            outputPath,
            ZipArchiveMode.Update);

        var item = workbook.GetEntry("customXml/item1.xml")
            ?? throw new Exception("Could not find customXml/item1.xml");

        XDocument xml;

        using (var stream = item.Open())
        {
            xml = XDocument.Load(stream);
        }

        var dataMashup = xml.Root
            ?? throw new Exception("Could not find DataMashup");

        byte[] mashupBytes =
            Convert.FromBase64String(dataMashup.Value);

        int oldZipLength =
            BitConverter.ToInt32(mashupBytes, 4);

        byte[] oldZipBytes = mashupBytes
            .Skip(8)
            .Take(oldZipLength)
            .ToArray();

        // Everything after the embedded ZIP must be preserved.
        byte[] trailingBytes = mashupBytes
            .Skip(8 + oldZipLength)
            .ToArray();

        // Make a writable copy of the embedded ZIP.
        using var zipStream = new MemoryStream();
        zipStream.Write(oldZipBytes, 0, oldZipBytes.Length);
        zipStream.Position = 0;

        using (var mashupZip =
            new ZipArchive(zipStream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var section = mashupZip.GetEntry("Formulas/Section1.m")
                ?? throw new Exception("Could not find Formulas/Section1.m");

            string mCode;

            using (var reader =
                new StreamReader(
                    section.Open(),
                    Encoding.UTF8,
                    detectEncodingFromByteOrderMarks: true))
            {
                mCode = reader.ReadToEnd();
            }

            // For now, replace our known test SQL.
            mCode = mCode.Replace(
                "select * from vehicle_trips",
                sql);

            // ZIP entries can't be overwritten directly.
            section.Delete();

            var newSection =
                mashupZip.CreateEntry("Formulas/Section1.m");

            using var writer =
                new StreamWriter(newSection.Open(), new UTF8Encoding(false));

            writer.Write(mCode);
        }

        byte[] newZipBytes = zipStream.ToArray();

        // Rebuild the complete DataMashup binary.
        using var newMashup = new MemoryStream();

        // Preserve bytes 0-3.
        newMashup.Write(mashupBytes, 0, 4);

        // Bytes 4-7 = new ZIP length.
        var lengthBytes = BitConverter.GetBytes(newZipBytes.Length);
        newMashup.Write(lengthBytes, 0, lengthBytes.Length);
        
        // New embedded ZIP.
        newMashup.Write(newZipBytes, 0, newZipBytes.Length);

        // Preserve everything after it.
        newMashup.Write(trailingBytes, 0, trailingBytes.Length);

        dataMashup.Value =
            Convert.ToBase64String(newMashup.ToArray());

        // Replace item1.xml in the XLSX.
        item.Delete();

        var newItem =
            workbook.CreateEntry("customXml/item1.xml");

        using var xmlStream = newItem.Open();

        xml.Save(xmlStream);
    }
}