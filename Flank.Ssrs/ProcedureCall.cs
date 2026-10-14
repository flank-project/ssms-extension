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
