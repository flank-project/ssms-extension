using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Flank.Excel;

public static partial class ExcelExporter
{
    private static readonly XNamespace Sheet =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // Template layout: A1 how-to note, A2:B2 last refreshed, row 3 blank,
    // the data table from row 4. Parameters go in rows 4.. and push the table
    // down, with a blank row between.
    private const int TemplateTableRow = 4;

    // With parameters: a header row (Parameter / User Input), then a row each
    // (an inline table takes its label's row, its own header row and rows,
    // and a blank row after).
    private const int ParameterHeaderRow = 4;
    private const int FirstParameterRow = 5;

    /// <summary>Where each parameter goes on the sheet, and the query's table below them.</summary>
    private sealed class SheetLayout
    {
        public SheetLayout(IReadOnlyList<int> rows, int tableRow)
        {
            Rows = rows;
            TableRow = tableRow;
        }

        /// <summary>
        /// Each parameter's row: its label in A, and its value in B; for an
        /// inline table, a hint in B, with the table (from column A) below.
        /// </summary>
        public IReadOnlyList<int> Rows { get; }

        /// <summary>The query's table's header row.</summary>
        public int TableRow { get; }
    }

    private static SheetLayout Layout(IReadOnlyList<WorkbookParameter> parameters)
    {
        if (parameters.Count == 0)
            return new SheetLayout(Array.Empty<int>(), TemplateTableRow);

        var rows = new List<int>();
        int row = FirstParameterRow;
        bool blankAfter = false;

        foreach (var parameter in parameters)
        {
            rows.Add(row);

            if (parameter.Table?.Layout == WorkbookTableLayout.Inline)
            {
                // Label, the table's header and rows, then a blank row, so
                // typing a value in the next parameter's cell doesn't extend
                // the table over it.
                row += 2 + parameter.Table.DataRowCount + 1;
                blankAfter = true;
            }
            else
            {
                row++;
                blankAfter = false;
            }
        }

        // A blank row before the query's table.
        return new SheetLayout(rows, blankAfter ? row : row + 1);
    }

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

        // Names for the cells and tables the query reads, and where they go.
        PrepareParameters(parameters, request.ParameterValues);
        var layout = Layout(parameters);

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

        SetNote(
            workbook,
            parameters.Any(p => p.Table != null) ? ParameterRefreshNote.Replace("in column B,", "in column B and rows in the tables,")
            : parameters.Count > 0 ? ParameterRefreshNote
            : RefreshNote);

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

            // Fitting the table's columns on refresh also narrows the parameter
            // cells above it (dates turn into ####).
            if (parameters.Count > 0)
                queryTable.Root.SetAttributeValue("adjustColumnWidth", "0");
        });

        // Likewise the "Last refreshed" time in A2, which would fit column A
        // (the parameter labels, and an inline table's first column) to it.
        if (parameters.Count > 0)
            UpdateXml(workbook, "xl/queryTables/queryTable2.xml", queryTable =>
                queryTable.Root!.SetAttributeValue("adjustColumnWidth", "0"));

        var cells = new CellWriter(workbook);

        if (parameters.Count > 0)
        {
            AddParameterCells(workbook, parameters, request.ParameterValues, layout, cells);
            AddParameterChecks(workbook, parameters, layout, cells);
            AddNotes(workbook, parameters, layout);

            if (parameters.Any(IsCheckbox))
                AddCheckboxFeature(workbook);

            var lists = Lists(parameters);
            if (lists.Count > 0)
                AddListParts(workbook, lists, cells);

            // After the lists' hidden sheet, whose names refer to it by position.
            AddTableSheets(workbook, parameters, cells);
        }

        if (columnNames != null && CanUseAsTableColumns(columnNames))
        {
            SetTableColumns(
                workbook,
                columnNames,
                request.Rows ?? Array.Empty<IReadOnlyList<object?>>(),
                tableRow: layout.TableRow,
                cells);
        }

        cells.SaveStyles();
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

        // Everything after the embedded ZIP must be preserved (lists add to its metadata).
        byte[] trailingBytes = mashupBytes
            .Skip(8 + oldZipLength)
            .ToArray();

        var lists = Lists(parameters);

        if (lists.Count > 0)
            trailingBytes = AddListMetadata(trailingBytes, lists);

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

            mCode = mCode.TrimEnd() + ListQueriesM(lists, server, database) + Environment.NewLine;

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
    /// Reads the parameter cells (each by its name, so inserting rows doesn't
    /// break it) and tables, checks each value, and passes them as a record
    /// (SQL parameter name -> value) to Value.NativeQuery, so values are sent
    /// as SQL parameters, never spliced into the SQL. A table's rows go as
    /// JSON text. Reading cells and the database in one query means Excel
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
            int? listNumber = Lists(parameters).Where(l => l.Parameter == parameters[i]).Select(l => (int?)l.Number).FirstOrDefault();
            string value = parameters[i].Table != null ? TableValueM(parameters[i]) : CheckedValueM(parameters[i], listNumber);
            checks.Append($"    Value{i} = {value},{Environment.NewLine}");
        }

        string values = string.Join(", ", Enumerable.Range(0, parameters.Count).Select(i => "Value" + i));
        string names = string.Join(", ", parameters.Select(p => "\"" + EscapeM(p.SqlName) + "\""));

        string query =
$@"shared Query1 = let
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
    /// M for one parameter's value, read from its named cell: blank is null
    /// (or an error if it's required); otherwise converted to its kind, with a
    /// message naming the parameter if it can't be. Data validation stops
    /// most bad values as they're typed, but not pasted ones.
    /// </summary>
    /// <param name="listNumber">For a parameter with a list: which one (List1 = 1).</param>
    internal static string CheckedValueM(WorkbookParameter parameter, int? listNumber = null)
    {
        // An empty named cell may come back as an empty table.
        string cell = $"try Excel.CurrentWorkbook(){{[Name=\"{EscapeM(parameter.CellName)}\"]}}[Content]{{0}}[Column1] otherwise null";
        string label = MText(parameter.DisplayName);

        return $"let v = {cell} in " +
               CheckM(parameter, label, parameter.Required ? " is required. Enter a value, then refresh." : null, listNumber);
    }

    /// <summary>
    /// M for a table-valued parameter's value: the named Excel table's
    /// non-blank rows, each value checked like a parameter's, as JSON text (an
    /// array of rows, each an array of column values in order); null for no
    /// rows, or an error if it's required.
    /// </summary>
    internal static string TableValueM(WorkbookParameter parameter)
    {
        var table = parameter.Table!;
        string nl = Environment.NewLine;
        string columns = string.Join(", ", table.Columns.Select(c => MText(c.DisplayName)));
        string blank = "(x) => x = null or (x is text and Text.Trim(x) = \"\")";

        // "Teams (row 2" & ", Team id)": the row counts entered rows, not sheet rows.
        string at = $"{MText(parameter.DisplayName + " (row ")} & Text.From(n + 1)";
        string checks = string.Join("," + nl, table.Columns.Select((c, i) =>
            $"                let v = row{{{i}}} in " +
            CheckM(c, $"at & {MText(", " + c.DisplayName + ")")}", c.Required ? " is required: fill it in, or clear the row." : null, null)));

        string unreadable = $"Couldn't read the {table.Name} table for {parameter.DisplayName}. " +
                            $"Its columns must be named {string.Join(", ", table.Columns.Select(c => c.DisplayName))}.";
        string ifEmpty = parameter.Required
            ? $"error Error.Record(\"Flank\", {MText(parameter.DisplayName + " needs at least one row. Enter one, then refresh.")})"
            : "null";

        return
$@"let
        Entered = try Table.SelectColumns(Excel.CurrentWorkbook(){{[Name=""{EscapeM(table.Name)}""]}}[Content], {{{columns}}})
            otherwise error Error.Record(""Flank"", {MText(unreadable)}),
        Rows = List.Select(Table.ToRows(Entered), (r) => not List.MatchesAll(r, {blank})),
        Checked = List.Transform(List.Positions(Rows), (n) => let row = Rows{{n}}, at = {at} in {{{nl}{checks}{nl}            }})
    in
        if List.IsEmpty(Rows) then {ifEmpty} else Text.FromBinary(Json.FromValue(Checked))";
    }

    /// <summary>An M text literal.</summary>
    private static string MText(string text) => "\"" + EscapeM(text) + "\"";

    /// <summary>
    /// M that checks v and converts it to the parameter's kind: blank is null,
    /// or an error if <paramref name="ifBlank"/> is given. Errors read
    /// "&lt;label&gt;&lt;message&gt;", e.g. "League must be a whole number.".
    /// </summary>
    /// <param name="label">M for the text that starts error messages: a literal, or an expression.</param>
    /// <param name="ifBlank">The message when it's blank and mustn't be, after the label; null if blank is fine.</param>
    private static string CheckM(WorkbookParameter parameter, string label, string? ifBlank, int? listNumber)
    {
        // One literal when the label is one, to keep the M readable.
        string Message(string text) =>
            label.Length >= 2 && label[0] == '"' && label[label.Length - 1] == '"' && label.IndexOf("\" & ", StringComparison.Ordinal) < 0
                ? label.Substring(0, label.Length - 1) + EscapeM(text) + "\""
                : $"{label} & {MText(text)}";

        string Fail(string text) => $"error Error.Record(\"Flank\", {Message(text)})";

        string whenBlank = ifBlank != null ? Fail(ifBlank) : "null";

        bool hasRange = parameter.Minimum != null && parameter.Maximum != null;
        string convert;

        switch (parameter.Kind)
        {
            case var _ when parameter.List != null && listNumber != null:
                convert = ListValueM(parameter, listNumber.Value, Fail);
                break;

            case WorkbookParameterKind.WholeNumber:
                string range = hasRange ? $" and n >= {parameter.Minimum} and n <= {parameter.Maximum}" : "";
                string between = hasRange
                    ? string.Format(CultureInfo.InvariantCulture, " between {0:N0} and {1:N0}", parameter.Minimum, parameter.Maximum)
                    : "";
                convert = $"try (let n = Number.From(v) in if n = Number.Round(n){range} then n else error \"\") " +
                          $"otherwise {Fail($" must be a whole number{between}.")}";
                break;

            case WorkbookParameterKind.Decimal:
                convert = $"try Number.From(v) otherwise {Fail(" must be a number.")}";
                break;

            case WorkbookParameterKind.Date:
                convert = $"try Date.From(v) otherwise {Fail(" must be a date.")}";
                break;

            case WorkbookParameterKind.DateTime:
                convert = $"try DateTime.From(v) otherwise {Fail(" must be a date and time.")}";
                break;

            case WorkbookParameterKind.TrueFalse:
                convert = $"try (if v is logical then v else Logical.From(v)) otherwise {Fail(" must be TRUE or FALSE.")}";
                break;

            case WorkbookParameterKind.Text when parameter.MaxLength != null:
                convert = $"let t = Text.From(v) in if Text.Length(t) <= {parameter.MaxLength} then t " +
                          $"else {Fail($" can be at most {parameter.MaxLength} characters.")}";
                break;

            default:
                convert = "v";
                break;
        }

        return $"if v = null or (v is text and Text.Trim(v) = \"\") then {whenBlank} else {convert}";
    }

    /// <summary>
    /// A bold header row (Parameter, User Input) in row 4, the parameter
    /// names in column A below it with their values in B, then the query
    /// table. An inline table goes below its label, from column A: a row
    /// added to a table shifts only the table's columns down, so starting in
    /// A keeps the labels and values below it together (for a table at least
    /// two columns wide). A table on its own sheet gets a link to it instead
    /// of a value. Each value cell gets a name (see <see cref="WorkbookParameter.CellName"/>),
    /// which the query reads it by.
    /// </summary>
    private static void AddParameterCells(
        ZipArchive workbook,
        IReadOnlyList<WorkbookParameter> parameters,
        IReadOnlyList<object?>? parameterValues,
        SheetLayout layout,
        CellWriter cells)
    {
        int tableRow = layout.TableRow;
        var inlineTables = new List<(WorkbookParameter Parameter, int Row)>();
        var links = new List<XElement>();
        int lastColumn = 2;

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

            var rows = new List<XElement>();

            var header = new XElement(Sheet + "row", new XAttribute("r", ParameterHeaderRow));
            string[] headings = { "Parameter", "User Input" };
            for (int c = 0; c < headings.Length; c++)
            {
                var heading = cells.Cell(ColumnLetter(c + 1) + ParameterHeaderRow, headings[c])!;
                heading.SetAttributeValue("s", cells.BoldStyle());
                header.Add(heading);
            }
            rows.Add(header);

            for (int i = 0; i < parameters.Count; i++)
            {
                var parameter = parameters[i];
                int r = layout.Rows[i];
                var row = new XElement(Sheet + "row",
                    new XAttribute("r", r),
                    cells.Cell("A" + r, parameter.DisplayName));
                rows.Add(row);

                var table = parameter.Table;

                if (table?.Layout == WorkbookTableLayout.Inline)
                {
                    // A hint on how to add rows, then the table's header row and rows.
                    var hint = cells.Cell("B" + r, InlineTableHint)!;
                    hint.SetAttributeValue("s", cells.HintStyle());
                    row.Add(hint);

                    var tableHeader = new XElement(Sheet + "row", new XAttribute("r", r + 1));
                    for (int c = 0; c < table.Columns.Count; c++)
                        tableHeader.Add(cells.Cell(ColumnLetter(c + 1) + (r + 1), table.Columns[c].DisplayName));
                    rows.Add(tableHeader);

                    for (int d = 0; d < table.DataRowCount; d++)
                    {
                        int dataRow = r + 2 + d;
                        var data = new XElement(Sheet + "row", new XAttribute("r", dataRow));
                        AddTableRowCells(data, table, d, dataRow, firstColumn: 1, cells);
                        rows.Add(data);
                    }

                    inlineTables.Add((parameter, r + 1));
                    lastColumn = Math.Max(lastColumn, table.Columns.Count);
                    continue;
                }

                if (table != null)
                {
                    // A link to the table's own sheet.
                    string text = $"Enter rows on the {table.SheetName} sheet";
                    var link = cells.Cell("B" + r, text)!;
                    link.SetAttributeValue("s", cells.LinkStyle());
                    row.Add(link);

                    links.Add(new XElement(Sheet + "hyperlink",
                        new XAttribute("ref", "B" + r),
                        new XAttribute("location", $"{SheetReference(table.SheetName)}!A1"),
                        new XAttribute("display", text)));
                    continue;
                }

                XElement? value = null;

                if (parameterValues != null && i < parameterValues.Count)
                {
                    // A list shows labels, so start with the label for the value.
                    value = cells.Cell("B" + r, ListLabelFor(parameter, parameterValues[i]), dateOnly: parameter.Kind == WorkbookParameterKind.Date);
                }

                if (IsCheckbox(parameter))
                {
                    // A blank checkbox shows greyed out: not set, so the blank rule applies.
                    value ??= new XElement(Sheet + "c", new XAttribute("r", "B" + r));
                    value.SetAttributeValue("s", cells.CheckboxStyle());
                }

                if (value != null)
                    row.Add(value);
            }

            // Rows must stay in order: insert before the (shifted) table rows.
            var firstTableRow = sheetData.Elements(Sheet + "row")
                .FirstOrDefault(r => (int)r.Attribute("r")! >= tableRow);

            if (firstTableRow != null)
                firstTableRow.AddBeforeSelf(rows);
            else
                sheetData.Add(rows);

            if (links.Count > 0)
            {
                // Schema order: ..., dataValidations, hyperlinks, printOptions, pageMargins, ...
                var after = sheet.Root.Element(Sheet + "pageMargins")
                    ?? throw new Exception("Could not find pageMargins");
                after.AddBeforeSelf(new XElement(Sheet + "hyperlinks", links));
            }

            sheet.Root.Element(Sheet + "dimension")?
                .SetAttributeValue("ref", $"A1:{ColumnLetter(lastColumn)}{tableRow + 1}");
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

            // Names move with their cells when rows are inserted, so the query keeps finding them.
            for (int i = 0; i < parameters.Count; i++)
            {
                if (parameters[i].Table == null)
                {
                    definedNames.Add(new XElement(Sheet + "definedName",
                        new XAttribute("name", parameters[i].CellName),
                        $"Query1!$B${layout.Rows[i]}"));
                }
            }
        });

        if (inlineTables.Count > 0)
        {
            var parts = inlineTables
                .Select(t => (t.Parameter.Table!, $"A{t.Row}:{ColumnLetter(t.Parameter.Table!.Columns.Count)}{t.Row + t.Parameter.Table!.DataRowCount}"))
                .ToList();

            AddTableParts(workbook, "xl/worksheets/sheet1.xml", parts);
        }
    }

    private const string InlineTableHint = "Enter rows in the table below. For more, right-click a row number in it > Insert";

    /// <summary>A sheet name as a formula or link refers to it: 'Team list'.</summary>
    private static string SheetReference(string sheetName) => "'" + sheetName.Replace("'", "''") + "'";

    /// <summary>One data row's cells for a table: its starting values, if it has any.</summary>
    private static void AddTableRowCells(XElement row, WorkbookTable table, int index, int rowNumber, int firstColumn, CellWriter cells)
    {
        if (index >= table.InitialRows.Count)
            return;

        var values = table.InitialRows[index];

        for (int c = 0; c < table.Columns.Count && c < values.Count; c++)
        {
            var cell = cells.Cell(ColumnLetter(firstColumn + c) + rowNumber, values[c], dateOnly: table.Columns[c].Kind == WorkbookParameterKind.Date);
            if (cell != null)
                row.Add(cell);
        }
    }

    /// <summary>
    /// Data validation on the value cells (type, range, length; TRUE/FALSE as a
    /// dropdown, for Excel without checkboxes), with the note as the message
    /// shown when a cell is selected, and required cells highlighted while
    /// they're empty. An inline table's columns get the same checks; a
    /// required column is highlighted in rows that have other values.
    /// </summary>
    private static void AddParameterChecks(
        ZipArchive workbook,
        IReadOnlyList<WorkbookParameter> parameters,
        SheetLayout layout,
        CellWriter cells)
    {
        var validations = new List<XElement>();
        var formats = new List<XElement>();
        var lists = Lists(parameters);
        int priority = 1;

        for (int i = 0; i < parameters.Count; i++)
        {
            var parameter = parameters[i];

            if (parameter.Table?.Layout == WorkbookTableLayout.Inline)
            {
                TableChecks(parameter.Table, layout.Rows[i] + 2, 1, cells, validations, formats, ref priority);
                continue;
            }

            if (parameter.Table != null)
                continue;

            string cell = "B" + layout.Rows[i];
            int? listNumber = lists.Where(l => l.Parameter == parameter).Select(l => (int?)l.Number).FirstOrDefault();

            var validation = DataValidation(parameter, cell, listNumber);
            if (validation != null)
                validations.Add(validation);

            if (parameter.Required)
                formats.Add(MissingValueFormat(cell, $"LEN(TRIM({cell}))=0", cells, ref priority));
        }

        AddChecks(workbook, "xl/worksheets/sheet1.xml", validations, formats);
    }

    /// <summary>The checks for a table's columns, whose first data row is <paramref name="firstRow"/>.</summary>
    private static void TableChecks(
        WorkbookTable table,
        int firstRow,
        int firstColumn,
        CellWriter cells,
        List<XElement> validations,
        List<XElement> formats,
        ref int priority)
    {
        int lastRow = firstRow + table.DataRowCount - 1;
        string first = ColumnLetter(firstColumn);
        string last = ColumnLetter(firstColumn + table.Columns.Count - 1);

        for (int c = 0; c < table.Columns.Count; c++)
        {
            var column = table.Columns[c];
            string letter = ColumnLetter(firstColumn + c);
            string range = $"{letter}{firstRow}:{letter}{lastRow}";

            var validation = DataValidation(column, range, null);
            if (validation != null)
                validations.Add(validation);

            // Relative to the range's first cell, like Excel's own rules.
            if (column.Required)
            {
                formats.Add(MissingValueFormat(
                    range,
                    $"AND(COUNTA(${first}{firstRow}:${last}{firstRow})>0,LEN(TRIM({letter}{firstRow}))=0)",
                    cells,
                    ref priority));
            }
        }
    }

    /// <summary>A light red fill on the cells while the formula is true.</summary>
    private static XElement MissingValueFormat(string range, string formula, CellWriter cells, ref int priority) =>
        new XElement(Sheet + "conditionalFormatting",
            new XAttribute("sqref", range),
            new XElement(Sheet + "cfRule",
                new XAttribute("type", "expression"),
                new XAttribute("dxfId", cells.MissingValueFormat()),
                new XAttribute("priority", priority++),
                new XElement(Sheet + "formula", formula)));

    /// <summary>Adds conditional formats and data validations to a sheet, in schema order.</summary>
    private static void AddChecks(ZipArchive workbook, string sheetPath, IReadOnlyList<XElement> validations, IReadOnlyList<XElement> formats)
    {
        if (validations.Count == 0 && formats.Count == 0)
            return;

        UpdateXml(workbook, sheetPath, sheet =>
        {
            // Schema order: sheetData, ..., mergeCells, phoneticPr, conditionalFormatting, dataValidations, ..., pageMargins.
            var sheetData = sheet.Root!.Element(Sheet + "sheetData")!;
            var after = sheetData.ElementsAfterSelf()
                .LastOrDefault(e => e.Name == Sheet + "mergeCells" || e.Name == Sheet + "phoneticPr")
                ?? sheetData;

            var added = new List<XElement>(formats);
            if (validations.Count > 0)
                added.Add(new XElement(Sheet + "dataValidations", new XAttribute("count", validations.Count), validations));

            after.AddAfterSelf(added.ToArray());
        });
    }

    private static XElement? DataValidation(WorkbookParameter parameter, string cell, int? listNumber)
    {
        string? type = null;
        string? comparison = null;
        string? formula1 = null;
        string? formula2 = null;
        string error = "";
        bool hasRange = parameter.Minimum != null && parameter.Maximum != null;

        switch (parameter.Kind)
        {
            case var _ when parameter.List != null && listNumber != null:
                type = "list";
                formula1 = ListRangeName(listNumber.Value);
                error = "Choose one of the choices in the list.";
                break;

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

    // A list's dropdown shows labels, so a bit with choices stays a dropdown.
    private static bool IsCheckbox(WorkbookParameter parameter) =>
        parameter.Kind == WorkbookParameterKind.TrueFalse && parameter.List == null;

    private static readonly XNamespace FeaturePropertyBag =
        "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag";

    /// <summary>
    /// The part that defines Microsoft 365's in-cell checkbox, as Excel saves it:
    /// a checkbox control, used by cell format complement 0 (see
    /// <see cref="CellWriter.CheckboxStyle"/>). Older Excel ignores it and shows
    /// TRUE/FALSE, with the dropdown from data validation.
    /// </summary>
    private static void AddCheckboxFeature(ZipArchive workbook)
    {
        const string path = "xl/featurePropertyBag/featurePropertyBag.xml";

        CreateXml(workbook, path, new XElement(FeaturePropertyBag + "FeaturePropertyBags",
            new XElement(FeaturePropertyBag + "bag", new XAttribute("type", "Checkbox")),
            new XElement(FeaturePropertyBag + "bag", new XAttribute("type", "XFControls"),
                new XElement(FeaturePropertyBag + "bagId", new XAttribute("k", "CellControl"), "0")),
            new XElement(FeaturePropertyBag + "bag", new XAttribute("type", "XFComplement"),
                new XElement(FeaturePropertyBag + "bagId", new XAttribute("k", "XFControls"), "1")),
            new XElement(FeaturePropertyBag + "bag", new XAttribute("type", "XFComplements"), new XAttribute("extRef", "XFComplementsMapperExtRef"),
                new XElement(FeaturePropertyBag + "a", new XAttribute("k", "MappedFeaturePropertyBags"),
                    new XElement(FeaturePropertyBag + "bagId", "2")))));

        UpdateXml(workbook, "xl/_rels/workbook.xml.rels", rels =>
        {
            rels.Root!.Add(new XElement(PackageRelationships + "Relationship",
                new XAttribute("Id", "rIdFlankCheckboxes"),
                new XAttribute("Type", "http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag"),
                new XAttribute("Target", "featurePropertyBag/featurePropertyBag.xml")));
        });

        UpdateXml(workbook, "[Content_Types].xml", types =>
        {
            types.Root!.Add(new XElement(ContentTypes + "Override",
                new XAttribute("PartName", "/" + path),
                new XAttribute("ContentType", "application/vnd.ms-excel.featurepropertybag+xml")));
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

        /// <param name="dateOnly">Show dates without a time (for date parameters).</param>
        public XElement? Cell(string reference, object? value, bool dateOnly = false)
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
                    return DateCell(cell, date, dateOnly);

                case DateTimeOffset offset:
                    return DateCell(cell, offset.DateTime, dateOnly);

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

        private XElement DateCell(XElement cell, DateTime date, bool dateOnly)
        {
            cell.SetAttributeValue("s", dateOnly ? DateOnlyStyle() : DateStyle());
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

        private int? boldStyle;

        /// <summary>A cell format in the workbook's default font, bold (for headings).</summary>
        public int BoldStyle()
        {
            if (boldStyle != null)
                return boldStyle.Value;

            var root = Styles().Root!;
            var fonts = root.Element(Sheet + "fonts") ?? throw new Exception("Could not find fonts");

            // The default font plus <b/>, which comes first in a font's elements.
            var bold = new XElement(fonts.Elements(Sheet + "font").First());
            bold.Element(Sheet + "b")?.Remove();
            bold.AddFirst(new XElement(Sheet + "b"));
            fonts.Add(bold);

            int fontId = fonts.Elements(Sheet + "font").Count() - 1;
            fonts.SetAttributeValue("count", fontId + 1);

            var cellXfs = root.Element(Sheet + "cellXfs") ?? throw new Exception("Could not find cellXfs");

            cellXfs.Add(new XElement(Sheet + "xf",
                new XAttribute("numFmtId", 0),
                new XAttribute("fontId", fontId),
                new XAttribute("fillId", 0),
                new XAttribute("borderId", 0),
                new XAttribute("xfId", 0),
                new XAttribute("applyFont", 1)));

            int index = cellXfs.Elements(Sheet + "xf").Count() - 1;
            cellXfs.SetAttributeValue("count", index + 1);

            boldStyle = index;
            return index;
        }

        private int? hintStyle;
        private int? linkStyle;

        /// <summary>Grey italic, for hints.</summary>
        public int HintStyle() => hintStyle ??= FontStyle(italic: true, underline: false, rgb: "FF7F7F7F");

        /// <summary>Blue underlined, for links.</summary>
        public int LinkStyle() => linkStyle ??= FontStyle(italic: false, underline: true, rgb: "FF0563C1");

        /// <summary>A cell format in the default font, changed as given.</summary>
        private int FontStyle(bool italic, bool underline, string rgb)
        {
            var root = Styles().Root!;
            var fonts = root.Element(Sheet + "fonts") ?? throw new Exception("Could not find fonts");

            // Font elements go in schema order: b, i, ..., u, ..., sz, color, name, ...
            var font = new XElement(fonts.Elements(Sheet + "font").First());
            font.Elements().Where(e => e.Name == Sheet + "i" || e.Name == Sheet + "u" || e.Name == Sheet + "color").Remove();

            var size = font.Element(Sheet + "sz");
            var color = new XElement(Sheet + "color", new XAttribute("rgb", rgb));
            if (size != null)
                size.AddAfterSelf(color);
            else
                font.AddFirst(color);

            if (underline)
                (size ?? color).AddBeforeSelf(new XElement(Sheet + "u"));
            if (italic)
                font.AddFirst(new XElement(Sheet + "i"));

            fonts.Add(font);
            int fontId = fonts.Elements(Sheet + "font").Count() - 1;
            fonts.SetAttributeValue("count", fontId + 1);

            var cellXfs = root.Element(Sheet + "cellXfs") ?? throw new Exception("Could not find cellXfs");

            cellXfs.Add(new XElement(Sheet + "xf",
                new XAttribute("numFmtId", 0),
                new XAttribute("fontId", fontId),
                new XAttribute("fillId", 0),
                new XAttribute("borderId", 0),
                new XAttribute("xfId", 0),
                new XAttribute("applyFont", 1)));

            int index = cellXfs.Elements(Sheet + "xf").Count() - 1;
            cellXfs.SetAttributeValue("count", index + 1);
            return index;
        }

        private int? checkboxStyle;

        /// <summary>
        /// A cell format that shows TRUE/FALSE as a checkbox in Microsoft 365:
        /// complement 0 from the feature property bag (<see cref="AddCheckboxFeature"/>).
        /// </summary>
        public int CheckboxStyle()
        {
            if (checkboxStyle != null)
                return checkboxStyle.Value;

            var cellXfs = Styles().Root!.Element(Sheet + "cellXfs")
                ?? throw new Exception("Could not find cellXfs");

            cellXfs.Add(new XElement(Sheet + "xf",
                new XAttribute("numFmtId", 0),
                new XAttribute("fontId", 0),
                new XAttribute("fillId", 0),
                new XAttribute("borderId", 0),
                new XAttribute("xfId", 0),
                new XElement(Sheet + "extLst",
                    new XElement(Sheet + "ext",
                        new XAttribute("uri", "{C7286773-470A-42A8-94C5-96B5CB345126}"),
                        new XAttribute(XNamespace.Xmlns + "xfpb", FeaturePropertyBag.NamespaceName),
                        new XElement(FeaturePropertyBag + "xfComplement", new XAttribute("i", 0))))));

            int index = cellXfs.Elements(Sheet + "xf").Count() - 1;
            cellXfs.SetAttributeValue("count", index + 1);

            checkboxStyle = index;
            return index;
        }

        private int? dateOnlyStyle;

        // Built-in number format 22 (m/d/yyyy h:mm).
        private int DateStyle() => dateStyle ??= NumberFormatStyle(22);

        // Built-in number format 14 (m/d/yyyy).
        private int DateOnlyStyle() => dateOnlyStyle ??= NumberFormatStyle(14);

        // A cell format using a built-in number format.
        private int NumberFormatStyle(int numberFormat)
        {
            var cellXfs = Styles().Root!.Element(Sheet + "cellXfs")
                ?? throw new Exception("Could not find cellXfs");

            cellXfs.Add(new XElement(Sheet + "xf",
                new XAttribute("numFmtId", numberFormat),
                new XAttribute("fontId", 0),
                new XAttribute("fillId", 0),
                new XAttribute("borderId", 0),
                new XAttribute("xfId", 0),
                new XAttribute("applyNumberFormat", 1)));

            int index = cellXfs.Elements(Sheet + "xf").Count() - 1;
            cellXfs.SetAttributeValue("count", index + 1);
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
