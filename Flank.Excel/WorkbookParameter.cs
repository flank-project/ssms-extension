namespace Flank.Excel;

/// <summary>What kind of value a parameter cell takes, for validation.</summary>
public enum WorkbookParameterKind
{
    /// <summary>Anything; no checks.</summary>
    Any,
    Text,
    WholeNumber,
    Decimal,
    Date,
    DateTime,
    /// <summary>TRUE or FALSE (a SQL bit).</summary>
    TrueFalse,
}

/// <summary>
/// One parameter cell: the procedure parameter it feeds, what the user sees,
/// and what values it takes. Excel checks values as they're typed (data
/// validation) and the query checks them again on refresh, since pasting
/// gets past data validation.
/// </summary>
public sealed class WorkbookParameter
{
    /// <param name="sqlName">The procedure's parameter name, without '@'.</param>
    /// <param name="displayName">The label in column A.</param>
    public WorkbookParameter(string sqlName, string displayName)
    {
        SqlName = sqlName;
        DisplayName = displayName;
    }

    public string SqlName { get; }

    public string DisplayName { get; }

    /// <summary>Help shown next to the cell and when it's selected; empty for none.</summary>
    public string Note { get; set; } = "";

    public WorkbookParameterKind Kind { get; set; } = WorkbookParameterKind.Any;

    /// <summary>Longest text allowed, in characters; null for no limit.</summary>
    public int? MaxLength { get; set; }

    /// <summary>Smallest and largest whole numbers allowed (e.g. the SQL type's range).</summary>
    public long? Minimum { get; set; }

    public long? Maximum { get; set; }

    /// <summary>Must be filled in: highlighted while empty, and refreshing fails without it.</summary>
    public bool Required { get; set; }

    /// <summary>A dropdown of choices from a query, instead of free entry; null for none.</summary>
    public WorkbookList? List { get; set; }
}
