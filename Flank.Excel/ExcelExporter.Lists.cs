using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Flank.Excel;

/// <summary>
/// Dropdown lists for parameter cells. Each list is a Power Query query
/// (List1, List2, ...) loaded to a table on a hidden "Lists" sheet, wired up
/// like the template's own loaded queries: a connection, a query table, a
/// table, a defined name, and an entry in the query package's metadata
/// saying where it loads. The cell's data validation points at the table's
/// label column through a defined name (FlankList1 = List1[label]), so the
/// dropdown follows the table as a refresh grows or shrinks it.
/// </summary>
public static partial class ExcelExporter
{
    private const string ListsSheetName = "DropdownOptions";

    // The template has two connections, tables and query tables (Query1 and
    // LastRefreshed); list n uses number n + 2 of each.
    private const int TemplateQueryCount = 2;

    private static readonly XNamespace Relationships =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static readonly XNamespace PackageRelationships =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    private static readonly XNamespace ContentTypes =
        "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>
    /// The parameters with lists, numbered from 1, each list's query named
    /// after its parameter (<see cref="WorkbookList.QueryName"/>).
    /// </summary>
    private static IReadOnlyList<(WorkbookParameter Parameter, int Number)> Lists(IReadOnlyList<WorkbookParameter> parameters)
    {
        var lists = parameters.Where(p => p.List != null).Select((p, i) => (p, i + 1)).ToList();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Query1", "LastRefreshed" };

        foreach (var (parameter, _) in lists)
        {
            string name = SafeName(parameter.DisplayName) + "_options";
            string unique = name;

            for (int n = 2; !used.Add(unique); n++)
                unique = name + "_" + n;

            parameter.List!.QueryName = unique;
        }

        return lists;
    }

    /// <summary>
    /// A name both Power Query and an Excel table accept: letters, digits and
    /// underscores, not starting with a digit ("Team name" -> Team_name).
    /// </summary>
    private static string SafeName(string text)
    {
        var name = new StringBuilder();

        foreach (char c in text.Trim())
            name.Append(char.IsLetterOrDigit(c) ? c : '_');

        string result = name.ToString().Trim('_');

        if (result.Length == 0)
            return "Parameter";

        return char.IsDigit(result[0]) ? "_" + result : result;
    }

    private static string ListRangeName(int number) => "FlankList" + number;

    /// <summary>"shared List1 = ..." for each list, to add to the query section.</summary>
    private static string ListQueriesM(IReadOnlyList<(WorkbookParameter Parameter, int Number)> lists, string server, string database)
    {
        var m = new StringBuilder();

        foreach (var (parameter, number) in lists)
        {
            var list = parameter.List!;
            string columns = string.Join(", ", list.Columns.Select(c => "\"" + EscapeM(c) + "\""));

            if (list.Sql == null)
            {
                // Fixed choices: the table itself, so refreshing needs no database.
                bool labelled = list.Columns.Count > 1;
                string rows = string.Join(", ", list.Rows.Select(row => labelled
                    ? $"{{{MLiteral(row.Label)}, {MLiteral(row.Value)}}}"
                    : $"{{{MLiteral(row.Value)}}}"));

                m.Append($@"

shared {list.QueryName} = #table({{{columns}}}, {{{rows}}});");
                continue;
            }

            m.Append($@"

shared {list.QueryName} = let
    Source = Sql.Database(""{EscapeM(server)}"", ""{EscapeM(database)}"", [Query=""{EscapeM(list.Sql)}""]),
    Choices = Table.SelectColumns(Source, {{{columns}}})
in
    Choices;");
        }

        return m.ToString();
    }

    /// <summary>An M literal for a fixed choice: a number, logical, date, datetime or text.</summary>
    private static string MLiteral(object? value)
    {
        switch (value)
        {
            case null:
                return "null";
            case bool flag:
                return flag ? "true" : "false";
            case DateTime date:
                return date.TimeOfDay == TimeSpan.Zero
                    ? $"#date({date.Year}, {date.Month}, {date.Day})"
                    : $"#datetime({date.Year}, {date.Month}, {date.Day}, {date.Hour}, {date.Minute}, {date.Second})";
            case string text:
                return "\"" + EscapeM(text) + "\"";
            case IFormattable number:
                return number.ToString(null, System.Globalization.CultureInfo.InvariantCulture);
            default:
                return "\"" + EscapeM(value.ToString() ?? "") + "\"";
        }
    }

    /// <summary>
    /// M for a list parameter's value: the chosen label must be in the list;
    /// the value passed is that row's value column. Reading the list query
    /// here makes it a dependency, so the lookup uses freshly loaded choices
    /// whatever order Excel refreshes the queries in.
    /// </summary>
    /// <param name="fail">M for an error with the given message after the parameter's name.</param>
    private static string ListValueM(WorkbookParameter parameter, int number, Func<string, string> fail)
    {
        var list = parameter.List!;

        return
            $"let match = Table.SelectRows({list.QueryName}, each Text.From(Record.Field(_, \"{EscapeM(list.LabelColumn)}\")) = Text.From(v)) in " +
            $"if Table.IsEmpty(match) then {fail(" must be one of the choices in its list.")} " +
            $"else Record.Field(match{{0}}, \"{EscapeM(list.ValueColumn)}\")";
    }

    /// <summary>The label to show for a starting value (the value from a preview run).</summary>
    private static object? ListLabelFor(WorkbookParameter parameter, object? value)
    {
        if (value == null || value is DBNull || parameter.List == null)
            return value;

        string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";

        foreach (var row in parameter.List.Rows)
        {
            if (Convert.ToString(row.Value, CultureInfo.InvariantCulture) == text)
                return row.Label;
        }

        return value;
    }

    /// <summary>
    /// Adds a metadata entry per list query to the query package's metadata,
    /// which sits after the package zip in the DataMashup: [permissions length]
    /// [permissions] [metadata length] [metadata] [the rest]. The metadata is
    /// [version] [XML length] [XML] [content length] [content].
    /// </summary>
    private static byte[] AddListMetadata(byte[] afterPackage, IReadOnlyList<(WorkbookParameter Parameter, int Number)> lists)
    {
        int position = 0;
        int permissionsLength = BitConverter.ToInt32(afterPackage, position);
        position += 4 + permissionsLength;

        int metadataLength = BitConverter.ToInt32(afterPackage, position);
        int metadataStart = position + 4;
        byte[] metadata = afterPackage.Skip(metadataStart).Take(metadataLength).ToArray();
        byte[] rest = afterPackage.Skip(metadataStart + metadataLength).ToArray();

        int xmlLength = BitConverter.ToInt32(metadata, 4);
        byte[] xmlBytes = metadata.Skip(8).Take(xmlLength).ToArray();
        byte[] afterXml = metadata.Skip(8 + xmlLength).ToArray();

        bool bom = xmlBytes.Length >= 3 && xmlBytes[0] == 0xEF && xmlBytes[1] == 0xBB && xmlBytes[2] == 0xBF;
        var xml = XDocument.Parse(Encoding.UTF8.GetString(xmlBytes, bom ? 3 : 0, xmlBytes.Length - (bom ? 3 : 0)));
        var items = xml.Root!.Element("Items")!;

        foreach (var (parameter, _) in lists)
        {
            string name = parameter.List!.QueryName;

            XElement Entry(string type, string value) =>
                new XElement("Entry", new XAttribute("Type", type), new XAttribute("Value", value));

            items.Add(
                new XElement("Item",
                    new XElement("ItemLocation",
                        new XElement("ItemType", "Formula"),
                        new XElement("ItemPath", "Section1/" + name)),
                    new XElement("StableEntries",
                        Entry("IsPrivate", "l0"),
                        Entry("QueryID", "s" + Guid.NewGuid()),
                        Entry("FillEnabled", "l1"),
                        Entry("FillObjectType", "sTable"),
                        Entry("FillToDataModelEnabled", "l0"),
                        Entry("FillTarget", "s" + name),
                        Entry("FillTargetNameCustomized", "l1"),
                        Entry("ResultType", "sTable"),
                        Entry("BufferNextRefresh", "l1"),
                        Entry("AddedToDataModel", "l0"))),
                new XElement("Item",
                    new XElement("ItemLocation",
                        new XElement("ItemType", "Formula"),
                        new XElement("ItemPath", "Section1/" + name + "/Source")),
                    new XElement("StableEntries")));
        }

        byte[] newXml = Encoding.UTF8.GetBytes(xml.Declaration + xml.Root.ToString(SaveOptions.DisableFormatting));
        if (bom)
            newXml = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(newXml).ToArray();

        var newMetadata = new MemoryStream();
        newMetadata.Write(metadata, 0, 4);
        newMetadata.Write(BitConverter.GetBytes(newXml.Length), 0, 4);
        newMetadata.Write(newXml, 0, newXml.Length);
        newMetadata.Write(afterXml, 0, afterXml.Length);
        byte[] newMetadataBytes = newMetadata.ToArray();

        var result = new MemoryStream();
        result.Write(afterPackage, 0, 4 + permissionsLength);
        result.Write(BitConverter.GetBytes(newMetadataBytes.Length), 0, 4);
        result.Write(newMetadataBytes, 0, newMetadataBytes.Length);
        result.Write(rest, 0, rest.Length);
        return result.ToArray();
    }

    /// <summary>
    /// The hidden Lists sheet, with a table per list (two columns apart),
    /// filled with the starting choices; and each table's query table,
    /// connection and names.
    /// </summary>
    private static void AddListParts(
        ZipArchive workbook,
        IReadOnlyList<(WorkbookParameter Parameter, int Number)> lists,
        CellWriter cells)
    {
        var sheetData = new XElement(Sheet + "sheetData");
        var rows = new SortedDictionary<int, XElement>();
        var tableParts = new XElement(Sheet + "tableParts", new XAttribute("count", lists.Count));
        var sheetRelationships = new XElement(PackageRelationships + "Relationships");
        var definedNames = new List<XElement>();
        int lastRow = 1;
        int lastColumn = 1;

        XElement Row(int r)
        {
            if (!rows.TryGetValue(r, out var row))
                rows[r] = row = new XElement(Sheet + "row", new XAttribute("r", r));
            return row;
        }

        foreach (var (parameter, number) in lists)
        {
            var list = parameter.List!;
            int id = TemplateQueryCount + number;
            var columns = list.Columns;
            int firstColumn = (number - 1) * 3 + 1;
            string first = ColumnLetter(firstColumn);
            string last = ColumnLetter(firstColumn + columns.Count - 1);
            int tableLastRow = 1 + Math.Max(list.Rows.Count, 1);
            string tableRef = $"{first}1:{last}{tableLastRow}";

            lastRow = Math.Max(lastRow, tableLastRow);
            lastColumn = Math.Max(lastColumn, firstColumn + columns.Count - 1);

            for (int c = 0; c < columns.Count; c++)
            {
                Row(1).Add(cells.Cell(ColumnLetter(firstColumn + c) + 1, columns[c]));
            }

            for (int r = 0; r < list.Rows.Count; r++)
            {
                var (value, label) = list.Rows[r];
                var values = columns.Count == 1 ? new[] { label } : new[] { label, value };

                for (int c = 0; c < values.Length; c++)
                {
                    var cell = cells.Cell(ColumnLetter(firstColumn + c) + (r + 2), values[c]);
                    if (cell != null)
                        Row(r + 2).Add(cell);
                }
            }

            string tableRelationship = "rId" + number;
            tableParts.Add(new XElement(Sheet + "tablePart", new XAttribute(Relationships + "id", tableRelationship)));
            sheetRelationships.Add(new XElement(PackageRelationships + "Relationship",
                new XAttribute("Id", tableRelationship),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"),
                new XAttribute("Target", $"../tables/table{id}.xml")));

            var tableColumns = new XElement(Sheet + "tableColumns", new XAttribute("count", columns.Count));
            var queryTableFields = new XElement(Sheet + "queryTableFields", new XAttribute("count", columns.Count));

            for (int c = 0; c < columns.Count; c++)
            {
                tableColumns.Add(new XElement(Sheet + "tableColumn",
                    new XAttribute("id", c + 1),
                    new XAttribute("uniqueName", c + 1),
                    new XAttribute("name", columns[c]),
                    new XAttribute("queryTableFieldId", c + 1)));

                queryTableFields.Add(new XElement(Sheet + "queryTableField",
                    new XAttribute("id", c + 1),
                    new XAttribute("name", columns[c]),
                    new XAttribute("tableColumnId", c + 1)));
            }

            string name = list.QueryName;

            CreateXml(workbook, $"xl/tables/table{id}.xml", new XElement(Sheet + "table",
                new XAttribute("id", id),
                new XAttribute("name", name),
                new XAttribute("displayName", name),
                new XAttribute("ref", tableRef),
                new XAttribute("tableType", "queryTable"),
                new XAttribute("totalsRowShown", "0"),
                new XElement(Sheet + "autoFilter", new XAttribute("ref", tableRef)),
                tableColumns,
                new XElement(Sheet + "tableStyleInfo",
                    new XAttribute("name", "TableStyleMedium2"),
                    new XAttribute("showFirstColumn", "0"),
                    new XAttribute("showLastColumn", "0"),
                    new XAttribute("showRowStripes", "1"),
                    new XAttribute("showColumnStripes", "0"))));

            CreateXml(workbook, $"xl/tables/_rels/table{id}.xml.rels", new XElement(PackageRelationships + "Relationships",
                new XElement(PackageRelationships + "Relationship",
                    new XAttribute("Id", "rId1"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"),
                    new XAttribute("Target", $"../queryTables/queryTable{id}.xml"))));

            CreateXml(workbook, $"xl/queryTables/queryTable{id}.xml", new XElement(Sheet + "queryTable",
                new XAttribute("name", "ExternalData_" + id),
                new XAttribute("connectionId", id),
                new XAttribute("autoFormatId", "16"),
                new XAttribute("applyNumberFormats", "0"),
                new XAttribute("applyBorderFormats", "0"),
                new XAttribute("applyFontFormats", "0"),
                new XAttribute("applyPatternFormats", "0"),
                new XAttribute("applyAlignmentFormats", "0"),
                new XAttribute("applyWidthHeightFormats", "0"),
                new XElement(Sheet + "queryTableRefresh",
                    new XAttribute("nextId", columns.Count + 1),
                    queryTableFields)));

            UpdateXml(workbook, "xl/connections.xml", connections =>
            {
                connections.Root!.Add(new XElement(Sheet + "connection",
                    new XAttribute("id", id),
                    new XAttribute("keepAlive", "1"),
                    new XAttribute("name", "Query - " + name),
                    new XAttribute("description", $"Connection to the '{name}' query in the workbook."),
                    new XAttribute("type", "5"),
                    new XAttribute("refreshedVersion", "8"),
                    new XAttribute("background", "1"),
                    new XAttribute("saveData", "1"),
                    new XElement(Sheet + "dbPr",
                        new XAttribute("connection",
                            $"Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=\"\""),
                        new XAttribute("command", $"SELECT * FROM [{name}]"))));
            });

            definedNames.Add(new XElement(Sheet + "definedName",
                new XAttribute("name", "ExternalData_" + id),
                new XAttribute("localSheetId", "1"),
                new XAttribute("hidden", "1"),
                $"{ListsSheetName}!${first}$1:${last}${tableLastRow}"));

            // Data validation can't use a structured reference, but a name can.
            definedNames.Add(new XElement(Sheet + "definedName",
                new XAttribute("name", ListRangeName(number)),
                $"{name}[{EscapeStructuredName(list.LabelColumn)}]"));
        }

        foreach (var row in rows.Values)
            sheetData.Add(row);

        CreateXml(workbook, "xl/worksheets/sheet2.xml", new XElement(Sheet + "worksheet",
            new XAttribute(XNamespace.Xmlns + "r", Relationships.NamespaceName),
            new XElement(Sheet + "dimension", new XAttribute("ref", $"A1:{ColumnLetter(lastColumn)}{lastRow}")),
            sheetData,
            new XElement(Sheet + "pageMargins",
                new XAttribute("left", "0.7"), new XAttribute("right", "0.7"),
                new XAttribute("top", "0.75"), new XAttribute("bottom", "0.75"),
                new XAttribute("header", "0.3"), new XAttribute("footer", "0.3")),
            tableParts));

        CreateXml(workbook, "xl/worksheets/_rels/sheet2.xml.rels", sheetRelationships);

        UpdateXml(workbook, "xl/_rels/workbook.xml.rels", rels =>
        {
            rels.Root!.Add(new XElement(PackageRelationships + "Relationship",
                new XAttribute("Id", "rIdFlankLists"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                new XAttribute("Target", "worksheets/sheet2.xml")));
        });

        UpdateXml(workbook, "xl/workbook.xml", book =>
        {
            var sheets = book.Root!.Element(Sheet + "sheets")!;
            int sheetId = sheets.Elements(Sheet + "sheet").Max(s => (int)s.Attribute("sheetId")!) + 1;

            sheets.Add(new XElement(Sheet + "sheet",
                new XAttribute("name", ListsSheetName),
                new XAttribute("sheetId", sheetId),
                new XAttribute("state", "hidden"),
                new XAttribute(Relationships + "id", "rIdFlankLists")));

            book.Root.Element(Sheet + "definedNames")!.Add(definedNames.ToArray());
        });

        UpdateXml(workbook, "[Content_Types].xml", types =>
        {
            void Override(string part, string type) =>
                types.Root!.Add(new XElement(ContentTypes + "Override",
                    new XAttribute("PartName", part),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml." + type)));

            Override("/xl/worksheets/sheet2.xml", "worksheet+xml");

            foreach (var (_, number) in lists)
            {
                int id = TemplateQueryCount + number;
                Override($"/xl/tables/table{id}.xml", "table+xml");
                Override($"/xl/queryTables/queryTable{id}.xml", "queryTable+xml");
            }
        });
    }

    // In a structured reference, ' [ ] # are escaped with '.
    private static string EscapeStructuredName(string column)
    {
        var escaped = new StringBuilder();

        foreach (char c in column)
        {
            if (c == '\'' || c == '[' || c == ']' || c == '#')
                escaped.Append('\'');
            escaped.Append(c);
        }

        return escaped.ToString();
    }

    private static void CreateXml(ZipArchive workbook, string path, XElement root)
    {
        var document = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);

        using var output = workbook.CreateEntry(path).Open();
        document.Save(output);
    }
}
