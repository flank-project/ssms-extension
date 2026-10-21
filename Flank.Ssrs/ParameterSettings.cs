using System.Collections.Generic;
using System.Linq;

namespace Flank.Ssrs
{
    /// <summary>What a report or workbook does when a parameter is left blank.</summary>
    public enum BlankValue
    {
        /// <summary>A value must be entered.</summary>
        Required,

        /// <summary>Leave the parameter out of the call, so the procedure's own default applies.</summary>
        UseProcedureDefault,

        /// <summary>Pass NULL.</summary>
        PassNull,
    }

    /// <summary>Where a dropdown's choices come from.</summary>
    public enum ListSource
    {
        /// <summary>A stored procedure that takes no parameters.</summary>
        Procedure,

        /// <summary>A query written on the guardrails page.</summary>
        Query,

        /// <summary>Values typed in; the report or workbook holds them, no SQL runs.</summary>
        Values,
    }

    /// <summary>
    /// How one input parameter of a stored procedure is presented to the
    /// people using the report or workbook, chosen on the parameter review page.
    /// </summary>
    public sealed class ParameterSettings
    {
        public ParameterSettings(SprocParameter parameter)
        {
            Parameter = parameter;
            DisplayName = parameter.Name.TrimStart('@');
            DefaultValue = parameter.DefaultValue;
            IfBlank = parameter.HasDefault ? BlankValue.UseProcedureDefault : BlankValue.Required;
        }

        public SprocParameter Parameter { get; }

        /// <summary>The name with '@', as the procedure declares it.</summary>
        public string SqlName => Parameter.Name;

        /// <summary>What people see instead of the parameter name.</summary>
        public string DisplayName { get; set; }

        /// <summary>Help for whoever fills it in; empty for none.</summary>
        public string Note { get; set; } = "";

        /// <summary>The value it starts with, as text; null to start blank.</summary>
        public string? DefaultValue { get; set; }

        public BlankValue IfBlank { get; set; }

        /// <summary>Where a dropdown's choices come from.</summary>
        public ListSource ListSource { get; set; }

        /// <summary>
        /// For <see cref="ListSource.Procedure"/>: a stored procedure
        /// ([schema].[name]) whose first result set gives the choices.
        /// </summary>
        public string? ListProcedure { get; set; }

        /// <summary>For <see cref="ListSource.Query"/>: a query whose first result set gives the choices.</summary>
        public string? ListQuery { get; set; }

        /// <summary>The column passed to the parameter.</summary>
        public string? ListValueColumn { get; set; }

        /// <summary>The column shown; may be the same as the value column.</summary>
        public string? ListLabelColumn { get; set; }

        /// <summary>
        /// The choices: value, then label. From a preview run of the procedure
        /// or query, or for <see cref="ListSource.Values"/>, the values themselves.
        /// </summary>
        public IReadOnlyList<KeyValuePair<object?, object?>> ListRows { get; set; } = new KeyValuePair<object?, object?>[0];

        /// <summary>The SQL that returns the choices; null for typed-in values (or no list).</summary>
        public string? ListSql =>
            ListSource == ListSource.Procedure && ListProcedure != null ? "EXEC " + ListProcedure
            : ListSource == ListSource.Query && !string.IsNullOrWhiteSpace(ListQuery) ? ListQuery
            : null;

        public bool HasList =>
            ListValueColumn != null && ListLabelColumn != null &&
            (ListSource == ListSource.Values ? ListRows.Count > 0 : ListSql != null);

        /// <summary>Settings as the procedure declares the parameters, before any review.</summary>
        public static IReadOnlyList<ParameterSettings> FromProcedure(IReadOnlyList<SprocParameter> parameters) =>
            parameters.Where(p => !p.IsOutput).Select(p => new ParameterSettings(p)).ToList();
    }
}
