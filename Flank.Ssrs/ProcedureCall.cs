using System.Collections.Generic;
using System.Linq;

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
                p.IsOutput ? $"{p.Name} = {p.Name} OUTPUT" : $"{p.Name} = {p.Name}"));

            return declare + "EXEC " + qualifiedName + (arguments.Length > 0 ? " " + arguments : "");
        }

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
            var omitWhenNull = new HashSet<string>(
                (settings ?? new ParameterSettings[0])
                    .Where(s => s.IfBlank == BlankValue.UseProcedureDefault)
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
            string passing = string.Join(", ", inputs.Select(p => $"{p.Name} = {p.Name}"));

            return
                $"DECLARE @flank_arguments nvarchar(max) = CONCAT(N'', {string.Join(", ", arguments)});\r\n" +
                $"DECLARE @flank_call nvarchar(max) = N'{Quote(declareOutputs)}EXEC {Quote(qualifiedName)}' + " +
                "CASE WHEN @flank_arguments = N'' THEN N'' ELSE N' ' + STUFF(@flank_arguments, 1, 2, N'') END;\r\n" +
                $"EXEC sp_executesql @flank_call, N'{Quote(definitions)}', {passing};";
        }

        private static string Quote(string text) => text.Replace("'", "''");

        // sp_executesql parameters can't be text, ntext or image.
        private static string ParameterTypeDeclaration(SprocParameter parameter)
        {
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
