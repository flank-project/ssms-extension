using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Flank.Tests.Support;

/// <summary>
/// Reads the Power Query (M) code back out of a generated workbook: the
/// DataMashup in customXml/item1.xml is base64 of [4 bytes][zip length][zip][...],
/// and the zip holds Formulas/Section1.m.
/// </summary>
internal static class MashupReader
{
    public static string ReadSection1(string workbookPath)
    {
        using var workbook = ZipFile.OpenRead(workbookPath);

        var item = workbook.GetEntry("customXml/item1.xml")
            ?? throw new InvalidOperationException("Workbook has no customXml/item1.xml.");

        XDocument xml;
        using (var stream = item.Open())
        {
            xml = XDocument.Load(stream);
        }

        byte[] mashup = Convert.FromBase64String(xml.Root!.Value);
        int zipLength = BitConverter.ToInt32(mashup, 4);

        using var zip = new ZipArchive(new MemoryStream(mashup, 8, zipLength), ZipArchiveMode.Read);
        var section = zip.GetEntry("Formulas/Section1.m")
            ?? throw new InvalidOperationException("Mashup has no Formulas/Section1.m.");

        using var reader = new StreamReader(section.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
