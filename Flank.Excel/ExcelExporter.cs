using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Flank.Excel;

public static class ExcelExporter
{
    /// <summary>Workbook-level name for the parameter cells (names in A, values in B).</summary>
    public const string ParametersName = "FlankParameters";

    private static readonly XNamespace Sheet =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static void Generate(
        string outputPath,
        string sql,
        string server,
        string database)
    {
        Generate(outputPath, sql, server, database, Array.Empty<string>());
    }

    /// <summary>
    /// Creates a refreshable workbook. With parameters, the sheet starts with
    /// one row per parameter (name in column A, value in column B, for the user
    /// to fill in), then a blank row, then the data table. The query passes the
    /// values to <paramref name="sql"/> as SQL parameters (@name), so
    /// <paramref name="sql"/> should reference each one, e.g.
    /// "EXEC [dbo].[proc] @League_Id = @League_Id".
    /// </summary>
    /// <param name="parameterNames">Parameter names without '@'.</param>
    /// <param name="columnNames">
    /// The query's result columns, if known. The table is created with them, so
    /// the first refresh fills it without restructuring it. Restructuring (e.g.
    /// replacing the template's placeholder column) shifts the cells above the
    /// table, including the parameter cells.
    /// </param>
    public static void Generate(
        string outputPath,
        string sql,
        string server,
        string database,
        IReadOnlyList<string> parameterNames,
        IReadOnlyList<string>? columnNames = null)
    {
        var assembly = typeof(ExcelExporter).Assembly;

        using var templateStream =
            assembly.GetManifestResourceStream("Flank.Excel.template.xlsx")
            ?? throw new Exception("Could not find embedded template.xlsx");

        using (var outputStream = File.Create(outputPath))
        {
            templateStream.CopyTo(outputStream);
        }

        using var workbook = ZipFile.Open(
            outputPath,
            ZipArchiveMode.Update);

        UpdateQuery(workbook, sql, server, database, parameterNames);

        if (parameterNames.Count > 0)
        {
            AddParameterCells(workbook, parameterNames);
        }

        if (columnNames != null && CanUseAsTableColumns(columnNames))
        {
            SetTableColumns(
                workbook,
                columnNames,
                tableRow: parameterNames.Count > 0 ? parameterNames.Count + 2 : 1);
        }
    }

    // Excel table columns need unique, non-empty names; otherwise leave the
    // placeholder and let the first refresh name them.
    private static bool CanUseAsTableColumns(IReadOnlyList<string> columnNames)
    {
        return columnNames.Count > 0 &&
            columnNames.All(n => !string.IsNullOrWhiteSpace(n)) &&
            columnNames.Distinct(StringComparer.OrdinalIgnoreCase).Count() == columnNames.Count;
    }

    /// <summary>
    /// Replaces the template's placeholder column with the given columns: the
    /// header cells, the table's columns, and the query table's fields.
    /// </summary>
    private static void SetTableColumns(
        ZipArchive workbook,
        IReadOnlyList<string> columnNames,
        int tableRow)
    {
        int count = columnNames.Count;
        string lastColumn = ColumnLetter(count);
        string tableRef = $"A{tableRow}:{lastColumn}{tableRow + 1}";

        UpdateXml(workbook, "xl/worksheets/sheet1.xml", sheet =>
        {
            var rows = sheet.Root!.Element(Sheet + "sheetData")!.Elements(Sheet + "row").ToList();
            var header = rows.Single(r => (int)r.Attribute("r")! == tableRow);
            var data = rows.Single(r => (int)r.Attribute("r")! == tableRow + 1);

            // Keep the template's data cell style for the empty first row.
            var dataStyle = data.Element(Sheet + "c")?.Attribute("s")?.Value;

            header.Attribute("spans")?.Remove();
            data.Attribute("spans")?.Remove();
            header.RemoveNodes();
            data.RemoveNodes();

            for (int i = 0; i < count; i++)
            {
                string column = ColumnLetter(i + 1);

                header.Add(new XElement(Sheet + "c",
                    new XAttribute("r", column + tableRow),
                    new XAttribute("t", "inlineStr"),
                    new XElement(Sheet + "is", new XElement(Sheet + "t", columnNames[i]))));

                var cell = new XElement(Sheet + "c", new XAttribute("r", column + (tableRow + 1)));
                if (dataStyle != null)
                    cell.SetAttributeValue("s", dataStyle);
                data.Add(cell);
            }

            // Parameter cells use columns A-B.
            sheet.Root.Element(Sheet + "dimension")?
                .SetAttributeValue("ref", $"A1:{ColumnLetter(Math.Max(count, 2))}{tableRow + 1}");
        });

        UpdateXml(workbook, "xl/tables/table1.xml", table =>
        {
            var root = table.Root!;
            root.SetAttributeValue("ref", tableRef);
            root.Element(Sheet + "autoFilter")?.SetAttributeValue("ref", tableRef);

            var tableColumns = root.Element(Sheet + "tableColumns")
                ?? throw new Exception("Could not find tableColumns");

            tableColumns.RemoveNodes();
            tableColumns.SetAttributeValue("count", count);

            for (int i = 1; i <= count; i++)
            {
                tableColumns.Add(new XElement(Sheet + "tableColumn",
                    new XAttribute("id", i),
                    new XAttribute("uniqueName", i),
                    new XAttribute("name", columnNames[i - 1]),
                    new XAttribute("queryTableFieldId", i)));
            }
        });

        UpdateXml(workbook, "xl/queryTables/queryTable1.xml", queryTable =>
        {
            var refresh = queryTable.Root!.Element(Sheet + "queryTableRefresh")
                ?? throw new Exception("Could not find queryTableRefresh");

            refresh.SetAttributeValue("nextId", count + 1);

            var fields = refresh.Element(Sheet + "queryTableFields")
                ?? throw new Exception("Could not find queryTableFields");

            fields.RemoveNodes();
            fields.SetAttributeValue("count", count);

            for (int i = 1; i <= count; i++)
            {
                fields.Add(new XElement(Sheet + "queryTableField",
                    new XAttribute("id", i),
                    new XAttribute("name", columnNames[i - 1]),
                    new XAttribute("tableColumnId", i)));
            }
        });

        UpdateXml(workbook, "xl/workbook.xml", book =>
        {
            foreach (var name in book.Root!.Descendants(Sheet + "definedName"))
            {
                if ((string)name.Attribute("name")! == "ExternalData_1")
                {
                    name.Value = $"Query1!$A${tableRow}:${lastColumn}${tableRow + 1}";
                }
            }
        });
    }

    // 1 -> A, 26 -> Z, 27 -> AA.
    private static string ColumnLetter(int number)
    {
        string letters = "";

        while (number > 0)
        {
            int remainder = (number - 1) % 26;
            letters = (char)('A' + remainder) + letters;
            number = (number - 1) / 26;
        }

        return letters;
    }

    private static void UpdateQuery(
        ZipArchive workbook,
        string sql,
        string server,
        string database,
        IReadOnlyList<string> parameterNames)
    {
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

            mCode = parameterNames.Count > 0
                ? BuildParameterizedQuery(sql, server, database)
                : ReplaceTemplateQuery(mCode, sql, server, database);

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

    private static string ReplaceTemplateQuery(
        string mCode,
        string sql,
        string server,
        string database)
    {
        // For now, replace our known test SQL.
        string oldServer =
            "fomf-sandbox-ssdb.database.windows.net";

        string oldDatabase =
            "fomf-ss-sandbox-db";
        // Need 4 double quotes to represent 2 double quotes in the M code.
        string oldSql =
            @"SELECT CAST(NULL AS varchar(1)) AS [Click [Data]] > """"Refresh All"""" to load data] WHERE 1 = 0";

        if (!mCode.Contains(oldServer))
        {
            throw new Exception("Could not find template server.");
        }

        if (!mCode.Contains(oldDatabase))
        {
            throw new Exception("Could not find template database.");
        }

        if (!mCode.Contains(oldSql))
        {
            throw new Exception("Could not find template SQL.");
        }

        return mCode
            .Replace(oldServer, EscapeM(server))
            .Replace(oldDatabase, EscapeM(database))
            .Replace(oldSql, EscapeM(sql));
    }

    /// <summary>
    /// Reads the parameter cells into a record (name -> value) and passes it to
    /// Value.NativeQuery, so values are sent as SQL parameters, never spliced
    /// into the SQL. Reading cells and the database in one query means Excel
    /// asks once about privacy levels; separate queries fail outright with a
    /// "rebuild this data combination" error.
    /// </summary>
    private static string BuildParameterizedQuery(
        string sql,
        string server,
        string database)
    {
        return
$@"section Section1;

shared Query1 = let
    Parameters = Excel.CurrentWorkbook(){{[Name=""{ParametersName}""]}}[Content],
    Values = Record.FromList(Parameters[Column2], Parameters[Column1]),
    Source = Value.NativeQuery(Sql.Database(""{EscapeM(server)}"", ""{EscapeM(database)}""), ""{EscapeM(sql)}"", Values)
in
    Source;
";
    }

    private static string EscapeM(string value)
    {
        return value.Replace("\"", "\"\"");
    }

    /// <summary>
    /// Puts the parameter names in A1:An, leaves row n+1 blank, and moves the
    /// query table from A1 to row n+2.
    /// </summary>
    private static void AddParameterCells(
        ZipArchive workbook,
        IReadOnlyList<string> parameterNames)
    {
        int count = parameterNames.Count;
        int tableRow = count + 2;

        UpdateXml(workbook, "xl/worksheets/sheet1.xml", sheet =>
        {
            var sheetData = sheet.Root!.Element(Sheet + "sheetData")
                ?? throw new Exception("Could not find sheetData");

            // The template's table occupies rows 1-2; shift them down.
            foreach (var row in sheetData.Elements(Sheet + "row"))
            {
                int r = (int)row.Attribute("r")! + count + 1;
                row.SetAttributeValue("r", r);

                foreach (var cell in row.Elements(Sheet + "c"))
                {
                    string column = new string(((string)cell.Attribute("r")!).TakeWhile(char.IsLetter).ToArray());
                    cell.SetAttributeValue("r", column + r);
                }
            }

            for (int i = count; i >= 1; i--)
            {
                sheetData.AddFirst(
                    new XElement(Sheet + "row",
                        new XAttribute("r", i),
                        new XElement(Sheet + "c",
                            new XAttribute("r", "A" + i),
                            new XAttribute("t", "inlineStr"),
                            new XElement(Sheet + "is",
                                new XElement(Sheet + "t", parameterNames[i - 1])))));
            }

            sheet.Root.Element(Sheet + "dimension")?
                .SetAttributeValue("ref", $"A1:B{tableRow + 1}");
        });

        UpdateXml(workbook, "xl/tables/table1.xml", table =>
        {
            string tableRef = $"A{tableRow}:A{tableRow + 1}";
            table.Root!.SetAttributeValue("ref", tableRef);
            table.Root.Element(Sheet + "autoFilter")?.SetAttributeValue("ref", tableRef);
        });

        UpdateXml(workbook, "xl/queryTables/queryTable1.xml", queryTable =>
        {
            // By default the table inserts and deletes cells as it grows,
            // which shifts the parameter cells sideways. Overwrite instead.
            queryTable.Root!.SetAttributeValue("growShrinkType", "overwriteClear");
        });

        UpdateXml(workbook, "xl/workbook.xml", book =>
        {
            var definedNames = book.Root!.Element(Sheet + "definedNames")
                ?? throw new Exception("Could not find definedNames");

            foreach (var name in definedNames.Elements(Sheet + "definedName"))
            {
                if ((string)name.Attribute("name")! == "ExternalData_1")
                {
                    name.Value = $"Query1!$A${tableRow}:$A${tableRow + 1}";
                }
            }

            definedNames.AddFirst(
                new XElement(Sheet + "definedName",
                    new XAttribute("name", ParametersName),
                    $"Query1!$A$1:$B${count}"));
        });
    }

    private static void UpdateXml(
        ZipArchive workbook,
        string path,
        Action<XDocument> update)
    {
        var entry = workbook.GetEntry(path)
            ?? throw new Exception($"Could not find {path}");

        XDocument xml;

        using (var stream = entry.Open())
        {
            xml = XDocument.Load(stream);
        }

        update(xml);

        entry.Delete();

        using var output = workbook.CreateEntry(path).Open();
        xml.Save(output);
    }
}
