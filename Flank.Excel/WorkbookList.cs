namespace Flank.Excel;

/// <summary>
/// The choices for a parameter cell, from a query (typically a stored
/// procedure). It's loaded to a table on a hidden "Lists" sheet by its own
/// Power Query query, so Refresh All updates it; the cell's dropdown shows
/// the label column, and the main query looks up the value for the chosen label.
/// </summary>
public sealed class WorkbookList
{
    /// <param name="sql">
    /// The query that returns the choices, e.g. "EXEC [dbo].[get_leagues]";
    /// null for fixed choices, which the workbook holds (<see cref="Rows"/>).
    /// </param>
    /// <param name="valueColumn">The column passed to the parameter.</param>
    /// <param name="labelColumn">The column shown in the dropdown; may be the same as the value column.</param>
    public WorkbookList(string? sql, string valueColumn, string labelColumn)
    {
        Sql = sql;
        ValueColumn = valueColumn;
        LabelColumn = labelColumn;
    }

    public string? Sql { get; }

    public string ValueColumn { get; }

    public string LabelColumn { get; }

    /// <summary>The choices to start with (from a preview run): value, then label.</summary>
    public IReadOnlyList<(object? Value, object? Label)> Rows { get; set; } = Array.Empty<(object?, object?)>();

    /// <summary>
    /// The name of its query and table, after the parameter it fills (e.g.
    /// League_options), as people see it under Data > Queries &amp; Connections.
    /// Set when the workbook is generated.
    /// </summary>
    internal string QueryName { get; set; } = "";

    /// <summary>The list's columns on the DropdownOptions sheet: the label first, then the value if it's different.</summary>
    internal IReadOnlyList<string> Columns =>
        string.Equals(LabelColumn, ValueColumn, StringComparison.Ordinal)
            ? new[] { LabelColumn }
            : new[] { LabelColumn, ValueColumn };
}
