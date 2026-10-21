using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Flank.Excel;

/// <summary>
/// Names for what the query reads (each parameter cell, each table), notes
/// on the parameter labels, and table-valued parameters' tables: inline in
/// the parameter list, or on sheets of their own.
/// </summary>
public static partial class ExcelExporter
{
    /// <summary>
    /// Names each single-cell parameter's value cell (Flank_League_Id) and
    /// each table (after its parameter: Teams), and each own-sheet table's
    /// sheet; numbers the tables' parts; and takes tables' starting rows from
    /// the starting values.
    /// </summary>
    private static void PrepareParameters(IReadOnlyList<WorkbookParameter> parameters, IReadOnlyList<object?>? values)
    {
        // Table names and defined names share one namespace.
        var lists = Lists(parameters);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Query1", "LastRefreshed" };
        foreach (var (parameter, _) in lists)
            names.Add(parameter.List!.QueryName);

        var sheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Query1", ListsSheetName };
        int tableId = TemplateQueryCount + lists.Count;

        for (int i = 0; i < parameters.Count; i++)
        {
            var parameter = parameters[i];
            var table = parameter.Table;

            if (table == null)
            {
                parameter.CellName = Unique("Flank_" + SafeName(parameter.SqlName), names);
                continue;
            }

            table.Name = Unique(TableName(parameter.DisplayName), names);
            table.Id = ++tableId;
            table.InitialRows = values != null && i < values.Count && values[i] is IReadOnlyList<IReadOnlyList<object?>> rows
                ? rows
                : Array.Empty<IReadOnlyList<object?>>();

            if (table.Layout == WorkbookTableLayout.Sheet)
                table.SheetName = UniqueSheetName(parameter.DisplayName, sheets);
        }
    }

    private static string Unique(string name, HashSet<string> used)
    {
        string unique = name;

        for (int n = 2; !used.Add(unique); n++)
            unique = name + "_" + n;

        return unique;
    }

    /// <summary>
    /// A table name from a label: <see cref="SafeName"/>, unless that reads
    /// as a cell reference (A1, R1C1, R, C), which Excel doesn't allow.
    /// </summary>
    private static string TableName(string label)
    {
        string name = SafeName(label);

        return Regex.IsMatch(name, @"^([A-Za-z]{1,3}[0-9]+|[RrCc]|[Rr][0-9]*[Cc][0-9]*)$")
            ? "Table_" + name
            : name;
    }

    /// <summary>A sheet name from a label: at most 31 characters, none of []:*?/\, unique.</summary>
    private static string UniqueSheetName(string label, HashSet<string> used)
    {
        string name = Regex.Replace(label, @"[\[\]:*?/\\]", " ").Trim().Trim('\'').Trim();
        if (name.Length == 0)
            name = "Rows";

        name = Cut(name, 31);
        string unique = name;

        for (int n = 2; !used.Add(unique); n++)
        {
            string suffix = $" ({n})";
            unique = Cut(name, 31 - suffix.Length) + suffix;
        }

        return unique;
    }

    private static string Cut(string text, int length) => text.Length <= length ? text : text.Substring(0, length).TrimEnd();

    /// <summary>
    /// Excel tables for <paramref name="tables"/> (their header row and data
    /// rows, given as a range) on the sheet at <paramref name="sheetPath"/>,
    /// which must already have their header cells.
    /// </summary>
    private static void AddTableParts(ZipArchive workbook, string sheetPath, IReadOnlyList<(WorkbookTable Table, string Range)> tables)
    {
        string sheetFile = sheetPath.Substring(sheetPath.LastIndexOf('/') + 1);
        string relsPath = $"xl/worksheets/_rels/{sheetFile}.rels";
        var relationships = new List<XElement>();
        var tableParts = new List<XElement>();

        foreach (var (table, range) in tables)
        {
            string relationship = "rIdFlankTable" + table.Id;

            CreateXml(workbook, $"xl/tables/table{table.Id}.xml", new XElement(Sheet + "table",
                new XAttribute("id", table.Id),
                new XAttribute("name", table.Name),
                new XAttribute("displayName", table.Name),
                new XAttribute("ref", range),
                new XAttribute("totalsRowShown", "0"),
                new XElement(Sheet + "tableColumns",
                    new XAttribute("count", table.Columns.Count),
                    table.Columns.Select((c, i) => new XElement(Sheet + "tableColumn",
                        new XAttribute("id", i + 1),
                        new XAttribute("name", c.DisplayName)))),
                new XElement(Sheet + "tableStyleInfo",
                    new XAttribute("name", "TableStyleLight9"),
                    new XAttribute("showFirstColumn", "0"),
                    new XAttribute("showLastColumn", "0"),
                    new XAttribute("showRowStripes", "1"),
                    new XAttribute("showColumnStripes", "0"))));

            relationships.Add(new XElement(PackageRelationships + "Relationship",
                new XAttribute("Id", relationship),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"),
                new XAttribute("Target", $"../tables/table{table.Id}.xml")));

            tableParts.Add(new XElement(Sheet + "tablePart", new XAttribute(Relationships + "id", relationship)));
        }

        if (workbook.GetEntry(relsPath) != null)
            UpdateXml(workbook, relsPath, rels => rels.Root!.Add(relationships));
        else
            CreateXml(workbook, relsPath, new XElement(PackageRelationships + "Relationships", relationships));

        UpdateXml(workbook, sheetPath, sheet =>
        {
            var parts = sheet.Root!.Element(Sheet + "tableParts");

            if (parts == null)
            {
                parts = new XElement(Sheet + "tableParts");
                var extensions = sheet.Root.Element(Sheet + "extLst");
                if (extensions != null)
                    extensions.AddBeforeSelf(parts);
                else
                    sheet.Root.Add(parts);
            }

            parts.Add(tableParts);
            parts.SetAttributeValue("count", parts.Elements(Sheet + "tablePart").Count());
        });

        UpdateXml(workbook, "[Content_Types].xml", types =>
        {
            foreach (var (table, _) in tables)
            {
                types.Root!.Add(new XElement(ContentTypes + "Override",
                    new XAttribute("PartName", $"/xl/tables/table{table.Id}.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")));
            }
        });
    }

    /// <summary>
    /// A sheet per table-valued parameter that has its own: the table from A1
    /// (header row frozen), with its starting rows and its columns' checks.
    /// </summary>
    private static void AddTableSheets(ZipArchive workbook, IReadOnlyList<WorkbookParameter> parameters, CellWriter cells)
    {
        var tables = parameters
            .Select(p => p.Table)
            .Where(t => t?.Layout == WorkbookTableLayout.Sheet)
            .Select(t => t!)
            .ToList();

        for (int k = 0; k < tables.Count; k++)
        {
            var table = tables[k];

            // sheet2.xml is the lists' hidden sheet, whether or not there is one.
            string file = $"sheet{3 + k}.xml";
            string path = "xl/worksheets/" + file;
            string relationship = "rIdFlankSheet" + (k + 1);
            int columnCount = table.Columns.Count;
            int lastRow = 1 + table.DataRowCount;
            string range = $"A1:{ColumnLetter(columnCount)}{lastRow}";

            var sheetData = new XElement(Sheet + "sheetData");
            var header = new XElement(Sheet + "row", new XAttribute("r", 1));
            for (int c = 0; c < columnCount; c++)
                header.Add(cells.Cell(ColumnLetter(c + 1) + 1, table.Columns[c].DisplayName));
            sheetData.Add(header);

            for (int d = 0; d < table.InitialRows.Count; d++)
            {
                var row = new XElement(Sheet + "row", new XAttribute("r", d + 2));
                AddTableRowCells(row, table, d, d + 2, firstColumn: 1, cells);
                sheetData.Add(row);
            }

            var validations = new List<XElement>();
            var formats = new List<XElement>();
            int priority = 1;
            TableChecks(table, 2, 1, cells, validations, formats, ref priority);

            CreateXml(workbook, path, new XElement(Sheet + "worksheet",
                new XAttribute(XNamespace.Xmlns + "r", Relationships.NamespaceName),
                new XElement(Sheet + "dimension", new XAttribute("ref", range)),
                new XElement(Sheet + "sheetViews",
                    new XElement(Sheet + "sheetView",
                        new XAttribute("workbookViewId", "0"),
                        new XElement(Sheet + "pane",
                            new XAttribute("ySplit", "1"),
                            new XAttribute("topLeftCell", "A2"),
                            new XAttribute("activePane", "bottomLeft"),
                            new XAttribute("state", "frozen")),
                        new XElement(Sheet + "selection",
                            new XAttribute("pane", "bottomLeft"),
                            new XAttribute("activeCell", "A2"),
                            new XAttribute("sqref", "A2")))),
                new XElement(Sheet + "sheetFormatPr", new XAttribute("defaultRowHeight", "14.25")),
                new XElement(Sheet + "cols",
                    table.Columns.Select((c, i) => new XElement(Sheet + "col",
                        new XAttribute("min", i + 1),
                        new XAttribute("max", i + 1),
                        new XAttribute("width", Math.Max(16, Math.Min(60, c.DisplayName.Length + 6))),
                        new XAttribute("customWidth", "1")))),
                sheetData,
                formats,
                validations.Count > 0
                    ? new XElement(Sheet + "dataValidations", new XAttribute("count", validations.Count), validations)
                    : null,
                new XElement(Sheet + "pageMargins",
                    new XAttribute("left", "0.7"), new XAttribute("right", "0.7"),
                    new XAttribute("top", "0.75"), new XAttribute("bottom", "0.75"),
                    new XAttribute("header", "0.3"), new XAttribute("footer", "0.3"))));

            AddTableParts(workbook, path, new[] { (table, range) });

            UpdateXml(workbook, "xl/_rels/workbook.xml.rels", rels =>
            {
                rels.Root!.Add(new XElement(PackageRelationships + "Relationship",
                    new XAttribute("Id", relationship),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                    new XAttribute("Target", "worksheets/" + file)));
            });

            UpdateXml(workbook, "xl/workbook.xml", book =>
            {
                var sheets = book.Root!.Element(Sheet + "sheets")!;
                int sheetId = sheets.Elements(Sheet + "sheet").Max(s => (int)s.Attribute("sheetId")!) + 1;

                sheets.Add(new XElement(Sheet + "sheet",
                    new XAttribute("name", table.SheetName),
                    new XAttribute("sheetId", sheetId),
                    new XAttribute(Relationships + "id", relationship)));
            });

            UpdateXml(workbook, "[Content_Types].xml", types =>
            {
                types.Root!.Add(new XElement(ContentTypes + "Override",
                    new XAttribute("PartName", "/" + path),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            });
        }
    }

    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace Office = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace ExcelVml = "urn:schemas-microsoft-com:office:excel";

    /// <summary>
    /// Each parameter's note as a cell note on its label (shown on hover, red
    /// triangle in the corner), which moves with the cell. Notes are a
    /// comments part plus a VML drawing that places their boxes.
    /// </summary>
    private static void AddNotes(ZipArchive workbook, IReadOnlyList<WorkbookParameter> parameters, SheetLayout layout)
    {
        var notes = parameters
            .Select((p, i) => (Row: layout.Rows[i], Text: p.Note.Trim()))
            .Where(n => n.Text.Length > 0)
            .ToList();

        if (notes.Count == 0)
            return;

        CreateXml(workbook, "xl/comments1.xml", new XElement(Sheet + "comments",
            new XElement(Sheet + "authors", new XElement(Sheet + "author", "Flank")),
            new XElement(Sheet + "commentList",
                notes.Select(n => new XElement(Sheet + "comment",
                    new XAttribute("ref", "A" + n.Row),
                    new XAttribute("authorId", "0"),
                    new XElement(Sheet + "text",
                        new XElement(Sheet + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), n.Text)))))));

        var shapes = notes.Select((n, k) =>
        {
            // Zero-based; the box sits to the right of the label, tall enough for the text.
            int row = n.Row - 1;
            int height = 3 + n.Text.Length / 60;

            return new XElement(Vml + "shape",
                new XAttribute("id", "_x0000_s" + (1025 + k)),
                new XAttribute("type", "#_x0000_t202"),
                new XAttribute("style", "position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:180pt;height:60pt;z-index:" + (k + 1) + ";visibility:hidden"),
                new XAttribute("fillcolor", "#ffffe1"),
                new XAttribute(Office + "insetmode", "auto"),
                new XElement(Vml + "fill", new XAttribute("color2", "#ffffe1")),
                new XElement(Vml + "shadow", new XAttribute("on", "t"), new XAttribute("color", "black"), new XAttribute("obscured", "t")),
                new XElement(Vml + "path", new XAttribute(Office + "connecttype", "none")),
                new XElement(Vml + "textbox", new XAttribute("style", "mso-direction-alt:auto"),
                    new XElement("div", new XAttribute("style", "text-align:left"))),
                new XElement(ExcelVml + "ClientData", new XAttribute("ObjectType", "Note"),
                    new XElement(ExcelVml + "Anchor", $"1, 15, {row}, 2, 2, 60, {row + height}, 4"),
                    new XElement(ExcelVml + "AutoFill", "False"),
                    new XElement(ExcelVml + "Row", row),
                    new XElement(ExcelVml + "Column", 0)));
        });

        var vml = new XElement("xml",
            new XAttribute(XNamespace.Xmlns + "v", Vml.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "o", Office.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "x", ExcelVml.NamespaceName),
            new XElement(Office + "shapelayout", new XAttribute(Vml + "ext", "edit"),
                new XElement(Office + "idmap", new XAttribute(Vml + "ext", "edit"), new XAttribute("data", "1"))),
            new XElement(Vml + "shapetype",
                new XAttribute("id", "_x0000_t202"),
                new XAttribute("coordsize", "21600,21600"),
                new XAttribute(Office + "spt", "202"),
                new XAttribute("path", "m,l,21600r21600,l21600,xe"),
                new XElement(Vml + "stroke", new XAttribute("joinstyle", "miter")),
                new XElement(Vml + "path", new XAttribute("gradientshapeok", "t"), new XAttribute(Office + "connecttype", "rect"))),
            shapes);

        using (var output = workbook.CreateEntry("xl/drawings/vmlDrawing1.vml").Open())
        using (var writer = new StreamWriter(output, new UTF8Encoding(false)))
        {
            writer.Write(vml.ToString(SaveOptions.DisableFormatting));
        }

        UpdateXml(workbook, "xl/worksheets/_rels/sheet1.xml.rels", rels =>
        {
            rels.Root!.Add(
                new XElement(PackageRelationships + "Relationship",
                    new XAttribute("Id", "rIdFlankNotes"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"),
                    new XAttribute("Target", "../comments1.xml")),
                new XElement(PackageRelationships + "Relationship",
                    new XAttribute("Id", "rIdFlankNoteBoxes"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"),
                    new XAttribute("Target", "../drawings/vmlDrawing1.vml")));
        });

        UpdateXml(workbook, "xl/worksheets/sheet1.xml", sheet =>
        {
            // Schema order: ..., drawing, legacyDrawing, ..., tableParts, extLst.
            var legacyDrawing = new XElement(Sheet + "legacyDrawing", new XAttribute(Relationships + "id", "rIdFlankNoteBoxes"));
            var next = sheet.Root!.Element(Sheet + "tableParts") ?? sheet.Root.Element(Sheet + "extLst");

            if (next != null)
                next.AddBeforeSelf(legacyDrawing);
            else
                sheet.Root.Add(legacyDrawing);
        });

        UpdateXml(workbook, "[Content_Types].xml", types =>
        {
            types.Root!.AddFirst(new XElement(ContentTypes + "Default",
                new XAttribute("Extension", "vml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.vmlDrawing")));
            types.Root.Add(new XElement(ContentTypes + "Override",
                new XAttribute("PartName", "/xl/comments1.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")));
        });
    }
}
