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

        /// <summary>
        /// A stored procedure ([schema].[name]) whose first result set gives the
        /// choices, as a dropdown; null for free entry. Its first column is the
        /// value passed, its second (if any) the label shown.
        /// </summary>
        public string? ListProcedure { get; set; }

        public string? ListValueColumn { get; set; }

        public string? ListLabelColumn { get; set; }

        /// <summary>The choices from a preview run of the list procedure: value, then label.</summary>
        public IReadOnlyList<KeyValuePair<object?, object?>> ListRows { get; set; } = new KeyValuePair<object?, object?>[0];

        public bool HasList => ListProcedure != null && ListValueColumn != null && ListLabelColumn != null;

        /// <summary>Settings as the procedure declares the parameters, before any review.</summary>
        public static IReadOnlyList<ParameterSettings> FromProcedure(IReadOnlyList<SprocParameter> parameters) =>
            parameters.Where(p => !p.IsOutput).Select(p => new ParameterSettings(p)).ToList();
    }
}
