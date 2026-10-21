namespace Flank.Excel;

/// <summary>Where a table-valued parameter's rows are entered.</summary>
public enum WorkbookTableLayout
{
    /// <summary>
    /// In the parameter list: an Excel table between the parameter above and
    /// the one below, starting with <see cref="WorkbookTable.Rows"/> empty rows.
    /// </summary>
    Inline,

    /// <summary>An Excel table on a sheet of its own, linked from the parameter list.</summary>
    Sheet,
}

/// <summary>
/// A table-valued parameter: an Excel table with a column per column of the
/// table type. Refreshing passes its non-blank rows to the query as JSON (an
/// array of rows, each an array of column values in order), which the query
/// reads with OPENJSON into a variable of the table type.
/// </summary>
public sealed class WorkbookTable
{
    /// <param name="columns">
    /// The table type's columns, in order: name (as both label and SQL name)
    /// and checks. <see cref="WorkbookParameter.Required"/> on a column means
    /// it can't be blank in a row that has other values.
    /// </param>
    public WorkbookTable(IReadOnlyList<WorkbookParameter> columns)
    {
        Columns = columns;
    }

    public IReadOnlyList<WorkbookParameter> Columns { get; }

    public WorkbookTableLayout Layout { get; set; }

    /// <summary>For <see cref="WorkbookTableLayout.Inline"/>: the rows it starts with. People can insert more.</summary>
    public int Rows { get; set; } = 3;

    /// <summary>The Excel table's name, which the query reads it by. Set when the workbook is generated.</summary>
    internal string Name { get; set; } = "";

    /// <summary>Its table part's number (xl/tables/table{Id}.xml). Set when the workbook is generated.</summary>
    internal int Id { get; set; }

    /// <summary>For <see cref="WorkbookTableLayout.Sheet"/>: its sheet's name. Set when the workbook is generated.</summary>
    internal string SheetName { get; set; } = "";

    /// <summary>The rows it starts filled in with (from a preview run). Set when the workbook is generated.</summary>
    internal IReadOnlyList<IReadOnlyList<object?>> InitialRows { get; set; } = Array.Empty<IReadOnlyList<object?>>();

    /// <summary>Data rows the table is created with: an Excel table has at least one.</summary>
    internal int DataRowCount =>
        Math.Max(Layout == WorkbookTableLayout.Inline ? Math.Max(Rows, 1) : 1, InitialRows.Count);
}
