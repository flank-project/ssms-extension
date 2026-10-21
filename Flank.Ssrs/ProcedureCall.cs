using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Flank.Ssrs
{
    /// <summary>
    /// Builds the T-SQL to call a stored procedure from tools that can only
    /// pass input parameters (Power Query, SSRS text queries).
    /// </summary>
    public static class ProcedureCall
    {
        /// <summary>
        /// "EXEC [dbo].[p] @a = @a", with OUTPUT parameters bound to declared
        /// variables: "DECLARE @o int; EXEC [dbo].[p] @a = @a, @o = @o OUTPUT".
        /// Input parameters stay as @name for the caller to supply.
        /// </summary>
        public static string BuildExec(string qualifiedName, IReadOnlyList<SprocParameter> parameters)
        {
            var outputs = parameters.Where(p => p.IsOutput).ToList();

            string declare = outputs.Count == 0
                ? ""
                : "DECLARE " + string.Join(", ", outputs.Select(p => $"{p.Name} {TypeDeclaration(p)}")) + "; ";

            string arguments = string.Join(", ", parameters.Select(p =>
                p.IsOutput ? $"{p.Name} = {p.Name} OUTPUT"
                : p.IsTable ? $"{p.Name} = {TableVariable(p)}"
                : $"{p.Name} = {p.Name}"));

            return FillTables(parameters) + declare + "EXEC " + qualifiedName + (arguments.Length > 0 ? " " + arguments : "");
        }

        /// <summary>
        /// Table-valued parameters can't be passed by Power Query or SSRS, so
        /// each is passed as JSON text under its own name (@Rows, an array of
        /// rows, each an array of column values in order; see
        /// <see cref="TableValues.ToJson"/>), and the call declares a variable of
        /// the table type, fills it from the JSON with OPENJSON, and passes that.
        /// NULL JSON is an empty table. Needs compatibility level 130 or later.
        /// </summary>
        private static string FillTables(IReadOnlyList<SprocParameter> parameters)
        {
            var fill = new StringBuilder();

            foreach (var table in parameters.Where(p => p.IsTable && !p.IsOutput))
            {
                string variable = TableVariable(table);
                fill.Append($"DECLARE {variable} {table.TableType};\r\n");

                if (table.Columns.Count == 0)
                    continue;

                var columns = table.Columns;
                string names = string.Join(", ", columns.Select(c => QuoteName(c.Name)));
                string read = string.Join(", ", columns.Select((c, i) => $"[c{i}]"));
                string with = string.Join(", ", columns.Select((c, i) => $"[c{i}] {TypeDeclaration(c)} '$[{i}]'"));

                fill.Append($"INSERT INTO {variable} ({names}) SELECT {read} FROM OPENJSON({table.Name}) WITH ({with});\r\n");
            }

            return fill.ToString();
        }

        /// <summary>The variable a table-valued parameter's rows are put in: @flank_table_Rows.</summary>
        private static string TableVariable(SprocParameter table) => "@flank_table_" + table.Name.TrimStart('@');

        private static string QuoteName(string name) => "[" + name.Replace("]", "]]") + "]";

        /// <summary>
        /// Like <see cref="BuildExec(string, IReadOnlyList{SprocParameter})"/>, but
        /// parameters whose settings say "use the procedure's default" when blank
        /// are left out of the call when their value is NULL, so the procedure's
        /// own default applies. The call is built at run time and run with
        /// sp_executesql; inputs are still passed as @name for the caller to
        /// supply, so the same text works in Power Query, SSRS and Power BI.
        /// </summary>
        public static string BuildExec(
            string qualifiedName,
            IReadOnlyList<SprocParameter> parameters,
            IReadOnlyList<ParameterSettings>? settings)
        {
            // A table-valued parameter is always passed: with no rows, it's an empty table.
            var omitWhenNull = new HashSet<string>(
                (settings ?? new ParameterSettings[0])
                    .Where(s => s.IfBlank == BlankValue.UseProcedureDefault && !s.Parameter.IsTable)
                    .Select(s => s.SqlName),
                System.StringComparer.OrdinalIgnoreCase);

            var inputs = parameters.Where(p => !p.IsOutput).ToList();

            if (!inputs.Any(p => omitWhenNull.Contains(p.Name)))
                return BuildExec(qualifiedName, parameters);

            var outputs = parameters.Where(p => p.IsOutput).ToList();

            // OUTPUT parameters are declared and bound inside the built call.
            string declareOutputs = outputs.Count == 0
                ? ""
                : "DECLARE " + string.Join(", ", outputs.Select(p => $"{p.Name} {TypeDeclaration(p)}")) + "; ";

            var arguments = inputs
                .Select(p => omitWhenNull.Contains(p.Name)
                    ? $"CASE WHEN {p.Name} IS NOT NULL THEN N', {p.Name} = {p.Name}' END"
                    : $"N', {p.Name} = {p.Name}'")
                .Concat(outputs.Select(p => $"N', {p.Name} = {p.Name} OUTPUT'"));

            string definitions = string.Join(", ", inputs.Select(p => $"{p.Name} {ParameterTypeDeclaration(p)}"));
            string passing = string.Join(", ", inputs.Select(p => $"{p.Name} = {(p.IsTable ? TableVariable(p) : p.Name)}"));

            return
                FillTables(parameters) +
                $"DECLARE @flank_arguments nvarchar(max) = CONCAT(N'', {string.Join(", ", arguments)});\r\n" +
                $"DECLARE @flank_call nvarchar(max) = N'{Quote(declareOutputs)}EXEC {Quote(qualifiedName)}' + " +
                "CASE WHEN @flank_arguments = N'' THEN N'' ELSE N' ' + STUFF(@flank_arguments, 1, 2, N'') END;\r\n" +
                $"EXEC sp_executesql @flank_call, N'{Quote(definitions)}', {passing};";
        }

        private static string Quote(string text) => text.Replace("'", "''");

        // sp_executesql parameters can't be text, ntext or image.
        private static string ParameterTypeDeclaration(SprocParameter parameter)
        {
            if (parameter.IsTable)
                return parameter.TableType + " READONLY";

            switch (parameter.SqlType.ToLowerInvariant())
            {
                case "text": return "varchar(max)";
                case "ntext": return "nvarchar(max)";
                case "image": return "varbinary(max)";
                default: return TypeDeclaration(parameter);
            }
        }

        /// <summary>The T-SQL type for a parameter, e.g. varchar(50), decimal(10, 2), nvarchar(max).</summary>
        public static string TypeDeclaration(SprocParameter parameter)
        {
            if (parameter.IsTable)
                return parameter.TableType!;

            string type = parameter.SqlType.ToLowerInvariant();

            switch (type)
            {
                case "char":
                case "varchar":
                case "binary":
                case "varbinary":
                    return $"{type}({(parameter.MaxLength < 0 ? "max" : parameter.MaxLength.ToString())})";

                case "nchar":
                case "nvarchar":
                    // max_length is in bytes; two per character.
                    return $"{type}({(parameter.MaxLength < 0 ? "max" : (parameter.MaxLength / 2).ToString())})";

                case "decimal":
                case "numeric":
                    return $"{type}({parameter.Precision}, {parameter.Scale})";

                case "datetime2":
                case "datetimeoffset":
                case "time":
                    return $"{type}({parameter.Scale})";

                default:
                    return type;
            }
        }
    }
}
