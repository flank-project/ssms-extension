namespace Flank.Excel;

/// <summary>What to put in a generated workbook.</summary>
public sealed class WorkbookRequest
{
    public WorkbookRequest(string sql, string server, string database)
    {
        Sql = sql;
        Server = server;
        Database = database;
    }

    /// <summary>
    /// The query. With parameters it must reference each one as @name, e.g.
    /// "EXEC [dbo].[proc] @League_Id = @League_Id".
    /// </summary>
    public string Sql { get; }

    public string Server { get; }

    public string Database { get; }

    /// <summary>Parameter names without '@'. Each gets a row: name in A, value in B.</summary>
    public IReadOnlyList<string> ParameterNames { get; set; } = Array.Empty<string>();

    /// <summary>Initial values for the parameter cells, in the same order. Null leaves a cell empty.</summary>
    public IReadOnlyList<object?>? ParameterValues { get; set; }

    /// <summary>
    /// The query's result columns, if known. The table is created with them, so
    /// the first refresh fills it without restructuring it. Restructuring (e.g.
    /// replacing the template's placeholder column) shifts the cells above the
    /// table, including the parameter cells.
    /// </summary>
    public IReadOnlyList<string>? ColumnNames { get; set; }

    /// <summary>
    /// Data to show before the first refresh (e.g. from a preview run), one list
    /// of values per row, in column order. Needs <see cref="ColumnNames"/>.
    /// </summary>
    public IReadOnlyList<IReadOnlyList<object?>>? Rows { get; set; }

    /// <summary>When <see cref="Rows"/> were fetched; shown as the "Last refreshed" time.</summary>
    public DateTime? RowsFetchedAt { get; set; }
}
