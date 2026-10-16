using System.Globalization;
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

    // Template layout: A1 how-to note, A2:B2 last refreshed, row 3 blank,
    // the data table from row 4. Parameters go in rows 4.. and push the table
    // down, with a blank row between.
    private const int TemplateTableRow = 4;
    private const int FirstParameterRow = 4;

    private const string RefreshNote =
        "To update: Data > Refresh All. The first time, Excel asks you to sign in to the database.";

    private const string ParameterRefreshNote =
        "To update: enter values in column B, then Data > Refresh All. The first time, Excel asks you to " +
        "sign in to the database and about privacy levels: tick \"Ignore Privacy Levels checks for this " +
        "file\", click Save, then save the workbook.";

    public static void Generate(
        string outputPath,
        string sql,
        string server,
        string database)
    {
        Generate(outputPath, sql, server, database, Array.Empty<string>());
    }

    /// <summary>
    /// Creates a refreshable workbook. The sheet starts with a how-to note and
    /// a "Last refreshed" time. With parameters, one row per parameter follows
    /// (name in column A, value in column B, for the user to fill in), then a
    /// blank row, then the data table. The query passes the
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
        Generate(outputPath, new WorkbookRequest(sql, server, database)
        {
            ParameterNames = parameterNames,
            ColumnNames = columnNames,
        });
    }

    public static void Generate(string outputPath, WorkbookRequest request)
    {
        var parameters = request.EffectiveParameters;
        var columnNames = request.ColumnNames;

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

        UpdateQuery(workbook, request.Sql, request.Server, request.Database, parameters);

        SetNote(workbook, parameters.Count > 0 ? ParameterRefreshNote : RefreshNote);

        if (request.RowsFetchedAt is DateTime fetched)
        {
            // Same text the LastRefreshed query writes.
            SetText(workbook, "A2", "Last refreshed: " + fetched.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture));
        }

        // By default the table inserts and deletes cells as it grows, which
        // shifts the cells above it (note, timestamp, parameters). Overwrite.
        UpdateXml(workbook, "xl/queryTables/queryTable1.xml", queryTable =>
        {
            queryTable.Root!.SetAttributeValue("growShrinkType", "overwriteClear");
        });

        var cells = new CellWriter(workbook);

        if (parameters.Count > 0)
        {
            AddParameterCells(workbook, parameters, request.ParameterValues, cells);
            AddParameterChecks(workbook, parameters, cells);
        }

        if (columnNames != null && CanUseAsTableColumns(columnNames))
        {
            SetTableColumns(
                workbook,
                columnNames,
                request.Rows ?? Array.Empty<IReadOnlyList<object?>>(),
                tableRow: TableRow(parameters.Count),
                cells);
        }

        cells.SaveStyles();
    }

    private static int TableRow(int parameterCount)
    {
        return parameterCount > 0
            ? FirstParameterRow + parameterCount + 1
            : TemplateTableRow;
    }

    private static void SetNote(ZipArchive workbook, string note)
    {
        SetText(workbook, "A1", note);
    }

    private static void SetText(ZipArchive workbook, string reference, string text)
    {
        UpdateXml(workbook, "xl/worksheets/sheet1.xml", sheet =>
        {
            var cell = sheet.Root!.Descendants(Sheet + "c").Single(c => (string)c.Attribute("r")! == reference);

            cell.RemoveNodes();
            cell.SetAttributeValue("t", "inlineStr");
            cell.Add(new XElement(Sheet + "is", new XElement(Sheet + "t", text)));
        });
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
        IReadOnlyList<IReadOnlyList<object?>> rows,
        int tableRow,
        CellWriter cells)
    {
        int count = columnNames.Count;
        string lastColumn = ColumnLetter(count);

        // A table always has at least one (possibly empty) data row.
        int lastRow = tableRow + Math.Max(rows.Count, 1);
        string tableRef = $"A{tableRow}:{lastColumn}{lastRow}";

        UpdateXml(workbook, "xl/worksheets/sheet1.xml", sheet =>
        {
            var sheetRows = sheet.Root!.Element(Sheet + "sheetData")!.Elements(Sheet + "row").ToList();
            var header = sheetRows.Single(r => (int)r.Attribute("r")! == tableRow);
            // The empty first data row may not exist in the sheet yet.
            var data = sheetRows.SingleOrDefault(r => (int)r.Attribute("r")! == tableRow + 1);

            if (data == null)
            {
                data = new XElement(Sheet + "row", new XAttribute("r", tableRow + 1));
                header.AddAfterSelf(data);
            }

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

                if (rows.Count == 0)
                {
                    var cell = new XElement(Sheet + "c", new XAttribute("r", column + (tableRow + 1)));
                    if (dataStyle != null)
                        cell.SetAttributeValue("s", dataStyle);
                    data.Add(cell);
                }
            }

            // Preview rows, shown until the first refresh replaces them.
            var previous = data;

            for (int r = 0; r < rows.Count; r++)
            {
                int rowNumber = tableRow + 1 + r;
                var row = r == 0 ? data : new XElement(Sheet + "row", new XAttribute("r", rowNumber));

                for (int i = 0; i < count && i < rows[r].Count; i++)
                {
                    var cell = cells.Cell(ColumnLetter(i + 1) + rowNumber, rows[r][i]);
                    if (cell != null)
                        row.Add(cell);
                }

                if (r > 0)
                {
                    previous.AddAfterSelf(row);
                    previous = row;
                }
            }

            // Parameter cells use columns A-B.
            sheet.Root.Element(Sheet + "dimension")?
                .SetAttributeValue("ref", $"A1:{ColumnLetter(Math.Max(count, 2))}{lastRow}");
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
                    name.Value = $"Query1!$A${tableRow}:${lastColumn}${lastRow}";
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
        IReadOnlyList<WorkbookParameter> parameters)
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

            mCode = parameters.Count > 0
                ? ReplaceWithParameterizedQuery(mCode, sql, server, database, parameters)
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
    /// Reads the parameter cells (column B, in order), checks each value, and
    /// passes them as a record (SQL parameter name -> value) to
    /// Value.NativeQuery, so values are sent as SQL parameters, never spliced
    /// into the SQL. Reading cells and the database in one query means Excel
    /// asks once about privacy levels; separate queries fail outright with a
    /// "rebuild this data combination" error.
    /// </summary>
    private static string ReplaceWithParameterizedQuery(
        string mCode,
        string sql,
        string server,
        string database,
        IReadOnlyList<WorkbookParameter> parameters)
    {
        // Replace only Query1; other queries (LastRefreshed) stay as they are.
        int from = mCode.IndexOf("shared Query1 = ", StringComparison.Ordinal);

        if (from < 0)
        {
            throw new Exception("Could not find the template's Query1.");
        }

        // Query1 ends where the next "shared" declaration's leading blank lines start.
        int end = mCode.IndexOf("\nshared ", from, StringComparison.Ordinal);

        while (end > from && char.IsWhiteSpace(mCode[end - 1]))
        {
            end--;
        }

        string rest = end < 0 ? Environment.NewLine : mCode.Substring(end);

        var checks = new StringBuilder();

        for (int i = 0; i < parameters.Count; i++)
        {
            checks.Append($"    Value{i} = {CheckedValueM(parameters[i], i)},{Environment.NewLine}");
        }

        string values = string.Join(", ", Enumerable.Range(0, parameters.Count).Select(i => "Value" + i));
        string names = string.Join(", ", parameters.Select(p => "\"" + EscapeM(p.SqlName) + "\""));

        string query =
$@"shared Query1 = let
    Cells = Excel.CurrentWorkbook(){{[Name=""{ParametersName}""]}}[Content][Column2],
{checks}    Values = Record.FromList({{{values}}}, {{{names}}}),
    Source = Value.NativeQuery(Sql.Database(""{EscapeM(server)}"", ""{EscapeM(database)}""), ""{EscapeM(sql)}"", Values)
in
    Source;";

        return mCode.Substring(0, from) + query + rest;
    }

    private static string EscapeM(string value)
    {
        return value.Replace("\"", "\"\"");
    }

    /// <summary>
    /// M for one parameter's value: blank is null (or an error if it's
    /// required); otherwise converted to its kind, with a message naming the
    /// cell if it can't be. Data validation stops most bad values as they're
    /// typed, but not pasted ones.
    /// </summary>
    internal static string CheckedValueM(WorkbookParameter parameter, int index)
    {
        string cell = "B" + (FirstParameterRow + index);
        string label = $"{parameter.DisplayName} (cell {cell})";

        string Fail(string message) => $"error Error.Record(\"Flank\", \"{EscapeM(message)}\")";

        string ifBlank = parameter.Required
            ? Fail($"{label} is required. Enter a value, then refresh.")
            : "null";

        bool hasRange = parameter.Minimum != null && parameter.Maximum != null;
        string convert;

        switch (parameter.Kind)
        {
            case WorkbookParameterKind.WholeNumber:
                string range = hasRange ? $" and n >= {parameter.Minimum} and n <= {parameter.Maximum}" : "";
                string between = hasRange
                    ? string.Format(CultureInfo.InvariantCulture, " between {0:N0} and {1:N0}", parameter.Minimum, parameter.Maximum)
                    : "";
                convert = $"try (let n = Number.From(v) in if n = Number.Round(n){range} then n else error \"\") " +
                          $"otherwise {Fail($"{label} must be a whole number{between}.")}";
                break;

            case WorkbookParameterKind.Decimal:
                convert = $"try Number.From(v) otherwise {Fail($"{label} must be a number.")}";
                break;

            case WorkbookParameterKind.Date:
                convert = $"try Date.From(v) otherwise {Fail($"{label} must be a date.")}";
                break;

            case WorkbookParameterKind.DateTime:
                convert = $"try DateTime.From(v) otherwise {Fail($"{label} must be a date and time.")}";
                break;

            case WorkbookParameterKind.TrueFalse:
                convert = $"try (if v is logical then v else Logical.From(v)) otherwise {Fail($"{label} must be TRUE or FALSE.")}";
                break;

            case WorkbookParameterKind.Text when parameter.MaxLength != null:
                convert = $"let t = Text.From(v) in if Text.Length(t) <= {parameter.MaxLength} then t " +
                          $"else {Fail($"{label} can be at most {parameter.MaxLength} characters.")}";
                break;

            default:
                convert = "v";
                break;
        }

        return $"let v = Cells{{{index}}} in if v = null or (v is text and Text.Trim(v) = \"\") then {ifBlank} else {convert}";
    }

    /// <summary>Column C: "Required." and the note, if there's anything to say.</summary>
    private static string NoteText(WorkbookParameter parameter)
    {
        string note = parameter.Note.Trim();
        return parameter.Required ? ("Required. " + note).Trim() : note;
    }

    /// <summary>
    /// Puts the parameter names in column A from row 4, leaves a blank row, and
    /// moves the query table below them.
    /// </summary>
    private static void AddParameterCells(
        ZipArchive workbook,
        IReadOnlyList<WorkbookParameter> parameters,
        IReadOnlyList<object?>? parameterValues,
        CellWriter cells)
    {
        int count = parameters.Count;
        int lastParameterRow = FirstParameterRow + count - 1;
        int tableRow = TableRow(count);

        UpdateXml(workbook, "xl/worksheets/sheet1.xml", sheet =>
        {
            var sheetData = sheet.Root!.Element(Sheet + "sheetData")
                ?? throw new Exception("Could not find sheetData");

            // Shift the template's table rows down past the parameters and a gap.
            foreach (var row in sheetData.Elements(Sheet + "row").Where(r => (int)r.Attribute("r")! >= TemplateTableRow))
            {
                int r = (int)row.Attribute("r")! + (tableRow - TemplateTableRow);
                row.SetAttributeValue("r", r);

                foreach (var cell in row.Elements(Sheet + "c"))
                {
                    string column = new string(((string)cell.Attribute("r")!).TakeWhile(char.IsLetter).ToArray());
                    cell.SetAttributeValue("r", column + r);
                }
            }

            // Rows must stay in order: insert before the (shifted) table rows.
            var firstTableRow = sheetData.Elements(Sheet + "row")
                .FirstOrDefault(r => (int)r.Attribute("r")! >= tableRow);

            for (int i = 0; i < count; i++)
            {
                int r = FirstParameterRow + i;
                var row = new XElement(Sheet + "row",
                    new XAttribute("r", r),
                    new XElement(Sheet + "c",
                        new XAttribute("r", "A" + r),
                        new XAttribute("t", "inlineStr"),
                        new XElement(Sheet + "is",
                            new XElement(Sheet + "t", parameters[i].DisplayName))));

                if (parameterValues != null && i < parameterValues.Count)
                {
                    var value = cells.Cell("B" + r, parameterValues[i]);
                    if (value != null)
                        row.Add(value);
                }

                string note = NoteText(parameters[i]);
                if (note.Length > 0)
                    row.Add(cells.Cell("C" + r, note));

                if (firstTableRow != null)
                    firstTableRow.AddBeforeSelf(row);
                else
                    sheetData.Add(row);
            }

            string lastColumn = parameters.Any(p => NoteText(p).Length > 0) ? "C" : "B";
            sheet.Root.Element(Sheet + "dimension")?
                .SetAttributeValue("ref", $"A1:{lastColumn}{tableRow + 1}");
        });

        UpdateXml(workbook, "xl/tables/table1.xml", table =>
        {
            string tableRef = $"A{tableRow}:A{tableRow + 1}";
            table.Root!.SetAttributeValue("ref", tableRef);
            table.Root.Element(Sheet + "autoFilter")?.SetAttributeValue("ref", tableRef);
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
                    $"Query1!$A${FirstParameterRow}:$B${lastParameterRow}"));
        });
    }

    /// <summary>
    /// Data validation on the value cells (type, range, length; TRUE/FALSE as a
    /// dropdown), with the note as the message shown when a cell is selected,
    /// and required cells highlighted while they're empty.
    /// </summary>
    private static void AddParameterChecks(
        ZipArchive workbook,
        IReadOnlyList<WorkbookParameter> parameters,
        CellWriter cells)
    {
        var validations = new XElement(Sheet + "dataValidations");
        var formats = new List<XElement>();
        int priority = 1;

        for (int i = 0; i < parameters.Count; i++)
        {
            var parameter = parameters[i];
            string cell = "B" + (FirstParameterRow + i);

            var validation = DataValidation(parameter, cell);
            if (validation != null)
                validations.Add(validation);

            if (parameter.Required)
            {
                formats.Add(new XElement(Sheet + "conditionalFormatting",
                    new XAttribute("sqref", cell),
                    new XElement(Sheet + "cfRule",
                        new XAttribute("type", "expression"),
                        new XAttribute("dxfId", cells.MissingValueFormat()),
                        new XAttribute("priority", priority++),
                        new XElement(Sheet + "formula", $"LEN(TRIM({cell}))=0"))));
            }
        }

        if (!validations.HasElements && formats.Count == 0)
            return;

        validations.SetAttributeValue("count", validations.Elements().Count());

        UpdateXml(workbook, "xl/worksheets/sheet1.xml", sheet =>
        {
            // Schema order: sheetData, ..., mergeCells, phoneticPr, conditionalFormatting, dataValidations, ..., pageMargins.
            var sheetData = sheet.Root!.Element(Sheet + "sheetData")!;
            var after = sheetData.ElementsAfterSelf()
                .LastOrDefault(e => e.Name == Sheet + "mergeCells" || e.Name == Sheet + "phoneticPr")
                ?? sheetData;

            var added = new List<XElement>(formats);
            if (validations.HasElements)
                added.Add(validations);

            after.AddAfterSelf(added.ToArray());
        });
    }

    private static XElement? DataValidation(WorkbookParameter parameter, string cell)
    {
        string? type = null;
        string? comparison = null;
        string? formula1 = null;
        string? formula2 = null;
        string error = "";
        bool hasRange = parameter.Minimum != null && parameter.Maximum != null;

        switch (parameter.Kind)
        {
            case WorkbookParameterKind.WholeNumber:
                type = "whole";
                comparison = "between";
                formula1 = (parameter.Minimum ?? -9007199254740991L).ToString(CultureInfo.InvariantCulture);
                formula2 = (parameter.Maximum ?? 9007199254740991L).ToString(CultureInfo.InvariantCulture);
                error = hasRange
                    ? string.Format(CultureInfo.InvariantCulture, "Enter a whole number between {0:N0} and {1:N0}.", parameter.Minimum, parameter.Maximum)
                    : "Enter a whole number.";
                break;

            case WorkbookParameterKind.Decimal:
                type = "decimal";
                comparison = "between";
                formula1 = "-1E+300";
                formula2 = "1E+300";
                error = "Enter a number.";
                break;

            case WorkbookParameterKind.Date:
            case WorkbookParameterKind.DateTime:
                type = "date";
                comparison = "greaterThanOrEqual";
                formula1 = "1";
                error = parameter.Kind == WorkbookParameterKind.Date ? "Enter a date." : "Enter a date and time.";
                break;

            case WorkbookParameterKind.TrueFalse:
                type = "list";
                formula1 = "\"TRUE,FALSE\"";
                error = "Choose TRUE or FALSE.";
                break;

            case WorkbookParameterKind.Text when parameter.MaxLength != null:
                type = "textLength";
                comparison = "lessThanOrEqual";
                formula1 = parameter.MaxLength.Value.ToString(CultureInfo.InvariantCulture);
                error = $"Enter at most {parameter.MaxLength} characters.";
                break;
        }

        string note = parameter.Note.Trim();

        if (type == null && note.Length == 0)
            return null;

        var validation = new XElement(Sheet + "dataValidation");

        if (type != null)
        {
            validation.SetAttributeValue("type", type);
            if (comparison != null)
                validation.SetAttributeValue("operator", comparison);
            validation.SetAttributeValue("allowBlank", "1");
            validation.SetAttributeValue("showErrorMessage", "1");
            validation.SetAttributeValue("errorTitle", Limit(parameter.DisplayName, 32));
            validation.SetAttributeValue("error", Limit(error, 255));
        }

        if (note.Length > 0)
        {
            // Shown when the cell is selected.
            validation.SetAttributeValue("showInputMessage", "1");
            validation.SetAttributeValue("promptTitle", Limit(parameter.DisplayName, 32));
            validation.SetAttributeValue("prompt", Limit(note, 255));
        }

        validation.SetAttributeValue("sqref", cell);

        if (formula1 != null)
            validation.Add(new XElement(Sheet + "formula1", formula1));
        if (formula2 != null)
            validation.Add(new XElement(Sheet + "formula2", formula2));

        return validation;
    }

    // Excel's limits: 32 characters for titles, 255 for messages.
    private static string Limit(string text, int length) =>
        text.Length <= length ? text : text.Substring(0, length - 1) + "…";

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

    /// <summary>
    /// Writes typed cells: numbers, booleans and dates as values (dates with a
    /// date format), everything else as text. Adds the date style on demand.
    /// </summary>
    private sealed class CellWriter
    {
        private readonly ZipArchive workbook;
        private XDocument? styles;
        private int? dateStyle;

        public CellWriter(ZipArchive workbook)
        {
            this.workbook = workbook;
        }

        public XElement? Cell(string reference, object? value)
        {
            var cell = new XElement(Sheet + "c", new XAttribute("r", reference));

            switch (value)
            {
                case null:
                case DBNull:
                    return null;

                case bool b:
                    cell.SetAttributeValue("t", "b");
                    cell.Add(new XElement(Sheet + "v", b ? "1" : "0"));
                    return cell;

                case byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal:
                    cell.Add(new XElement(Sheet + "v", Convert.ToString(value, CultureInfo.InvariantCulture)));
                    return cell;

                case DateTime date:
                    return DateCell(cell, date);

                case DateTimeOffset offset:
                    return DateCell(cell, offset.DateTime);

                case byte[] bytes:
                    return TextCell(cell, "0x" + BitConverter.ToString(bytes).Replace("-", ""));

                default:
                    return TextCell(cell, Convert.ToString(value, CultureInfo.InvariantCulture) ?? "");
            }
        }

        private static XElement TextCell(XElement cell, string text)
        {
            cell.SetAttributeValue("t", "inlineStr");
            var t = new XElement(Sheet + "t", text);

            // Keep leading/trailing spaces.
            if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
                t.SetAttributeValue(XNamespace.Xml + "space", "preserve");

            cell.Add(new XElement(Sheet + "is", t));
            return cell;
        }

        private XElement DateCell(XElement cell, DateTime date)
        {
            cell.SetAttributeValue("s", DateStyle());
            cell.Add(new XElement(Sheet + "v", date.ToOADate().ToString("R", CultureInfo.InvariantCulture)));
            return cell;
        }

        private int? missingValueFormat;

        /// <summary>A differential format (for conditional formatting): a light red fill.</summary>
        public int MissingValueFormat()
        {
            if (missingValueFormat != null)
                return missingValueFormat.Value;

            var root = Styles().Root!;
            var dxfs = root.Element(Sheet + "dxfs");

            if (dxfs == null)
            {
                dxfs = new XElement(Sheet + "dxfs");
                (root.Element(Sheet + "cellStyles") ?? root.Element(Sheet + "cellXfs")!).AddAfterSelf(dxfs);
            }

            dxfs.Add(new XElement(Sheet + "dxf",
                new XElement(Sheet + "fill",
                    new XElement(Sheet + "patternFill",
                        new XElement(Sheet + "bgColor", new XAttribute("rgb", "FFFFC7CE"))))));

            int index = dxfs.Elements(Sheet + "dxf").Count() - 1;
            dxfs.SetAttributeValue("count", index + 1);

            missingValueFormat = index;
            return index;
        }

        private XDocument Styles()
        {
            if (styles != null)
                return styles;

            var entry = workbook.GetEntry("xl/styles.xml")
                ?? throw new Exception("Could not find xl/styles.xml");

            using (var stream = entry.Open())
            {
                styles = XDocument.Load(stream);
            }

            return styles;
        }

        // A cell format using built-in number format 22 (m/d/yyyy h:mm).
        private int DateStyle()
        {
            if (dateStyle != null)
                return dateStyle.Value;

            var cellXfs = Styles().Root!.Element(Sheet + "cellXfs")
                ?? throw new Exception("Could not find cellXfs");

            cellXfs.Add(new XElement(Sheet + "xf",
                new XAttribute("numFmtId", 22),
                new XAttribute("fontId", 0),
                new XAttribute("fillId", 0),
                new XAttribute("borderId", 0),
                new XAttribute("xfId", 0),
                new XAttribute("applyNumberFormat", 1)));

            int index = cellXfs.Elements(Sheet + "xf").Count() - 1;
            cellXfs.SetAttributeValue("count", index + 1);

            dateStyle = index;
            return index;
        }

        public void SaveStyles()
        {
            if (styles == null)
                return;

            workbook.GetEntry("xl/styles.xml")!.Delete();

            using var output = workbook.CreateEntry("xl/styles.xml").Open();
            styles.Save(output);
        }
    }
}
