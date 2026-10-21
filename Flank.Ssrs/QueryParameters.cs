using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Flank.Ssrs
{
    /// <summary>
    /// Parameters in a query from the query editor, written the T-SQL way: a
    /// @name the query uses but doesn't DECLARE (WHERE league_id = @League_Id).
    /// </summary>
    public static class QueryParameters
    {
        /// <summary>A type to fall back on when SQL Server can't suggest one.</summary>
        public const string DefaultType = "nvarchar(100)";

        /// <summary>Types offered in the type picker, most common first.</summary>
        public static readonly IReadOnlyList<string> CommonTypes = new[]
        {
            "int", "bigint", "smallint", "tinyint", "bit",
            "decimal(18, 2)", "float", "money",
            "date", "datetime2", "datetime",
            "nvarchar(50)", "nvarchar(100)", "nvarchar(4000)", "nvarchar(max)",
            "varchar(50)", "varchar(100)", "varchar(max)",
            "uniqueidentifier",
        };

        /// <summary>
        /// The undeclared @names the query uses, in order of first use (the
        /// casing of the first use). Empty if the query can't be parsed.
        /// </summary>
        public static IReadOnlyList<string> Find(string sql)
        {
            var parser = new TSql160Parser(initialQuotedIdentifiers: true);
            TSqlFragment fragment;

            using (var reader = new StringReader(sql ?? ""))
            {
                fragment = parser.Parse(reader, out IList<ParseError> errors);

                if (errors.Count > 0)
                    return Array.Empty<string>();
            }

            var visitor = new VariableVisitor();
            fragment.Accept(visitor);

            return visitor.Used
                .Where(name => !visitor.Declared.Contains(name))
                .ToList();
        }

        /// <summary>Why the query can't be parsed (line and message), or null if it can.</summary>
        public static string? ParseProblem(string sql)
        {
            var parser = new TSql160Parser(initialQuotedIdentifiers: true);

            using (var reader = new StringReader(sql ?? ""))
            {
                parser.Parse(reader, out IList<ParseError> errors);

                return errors.Count == 0 ? null : $"Line {errors[0].Line}: {errors[0].Message}";
            }
        }

        /// <summary>
        /// SQL Server's suggested type for each undeclared parameter
        /// (sp_describe_undeclared_parameters), by name with '@'. Empty if it
        /// can't suggest (e.g. a parameter used too loosely to tell).
        /// </summary>
        public static IReadOnlyDictionary<string, string> SuggestTypes(string sql, IDbConnection connection)
        {
            var types = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "sp_describe_undeclared_parameters";
                    command.CommandType = CommandType.StoredProcedure;

                    var tsql = command.CreateParameter();
                    tsql.ParameterName = "@tsql";
                    tsql.DbType = DbType.String;
                    tsql.Size = -1;
                    tsql.Value = sql;
                    command.Parameters.Add(tsql);

                    using (var reader = command.ExecuteReader())
                    {
                        int name = reader.GetOrdinal("name");
                        int type = reader.GetOrdinal("suggested_system_type_name");

                        while (reader.Read())
                            types[reader.GetString(name)] = reader.GetString(type);
                    }
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                // Suggestions are a convenience: the user picks types anyway.
                System.Diagnostics.Debug.WriteLine($"Flank: no type suggestions: {ex.Message}");
            }

            return types;
        }

        /// <summary>
        /// A parameter of the given T-SQL type ("int", "nvarchar(50)",
        /// "decimal(10, 2)", "varchar(max)"), with metadata as sys.parameters
        /// would give it. Null if the type isn't recognised.
        /// </summary>
        public static SprocParameter? Parameter(string name, string type)
        {
            var match = Regex.Match(
                (type ?? "").Trim(),
                @"^(?<type>[a-z][a-z0-9_]*)\s*(\(\s*(?<a>max|\d+)\s*(,\s*(?<b>\d+)\s*)?\))?$",
                RegexOptions.IgnoreCase);

            if (!match.Success)
                return null;

            string sqlType = match.Groups["type"].Value.ToLowerInvariant();
            string a = match.Groups["a"].Value;
            string b = match.Groups["b"].Value;
            bool max = a.Equals("max", StringComparison.OrdinalIgnoreCase);
            int length = !max && a.Length > 0 ? int.Parse(a, CultureInfo.InvariantCulture) : 0;

            var parameter = new SprocParameter { Name = "@" + name.TrimStart('@'), SqlType = sqlType };

            switch (sqlType)
            {
                case "char":
                case "varchar":
                case "binary":
                case "varbinary":
                    parameter.MaxLength = (short)(max ? -1 : Math.Max(length, 1));
                    return parameter;

                case "nchar":
                case "nvarchar":
                    // max_length is in bytes: two per character.
                    parameter.MaxLength = (short)(max ? -1 : Math.Max(length, 1) * 2);
                    return parameter;

                case "decimal":
                case "numeric":
                    parameter.Precision = (byte)(a.Length > 0 && !max ? length : 18);
                    parameter.Scale = (byte)(b.Length > 0 ? int.Parse(b, CultureInfo.InvariantCulture) : 0);
                    return parameter;

                case "datetime2":
                case "datetimeoffset":
                case "time":
                    parameter.Scale = (byte)(a.Length > 0 && !max ? length : 7);
                    return parameter;

                case "int":
                case "bigint":
                case "smallint":
                case "tinyint":
                case "bit":
                case "float":
                case "real":
                case "money":
                case "smallmoney":
                case "date":
                case "datetime":
                case "smalldatetime":
                case "uniqueidentifier":
                    return a.Length == 0 ? parameter : null;

                default:
                    return null;
            }
        }

        /// <summary>
        /// CREATE PROCEDURE for a query: its parameters become the
        /// procedure's (no defaults), and the query is the body.
        /// </summary>
        public static string CreateProcedureScript(string qualifiedName, IReadOnlyList<SprocParameter> parameters, string sql)
        {
            string declarations = parameters.Count == 0
                ? "\n"
                : "\n" + string.Join(",\n", parameters.Select(p => $"    {p.Name} {ProcedureCall.TypeDeclaration(p)}")) + "\n";

            return
                $"CREATE PROCEDURE {qualifiedName}{declarations}AS\nBEGIN\n    SET NOCOUNT ON;\n\n" +
                sql.Trim().TrimEnd(';') + ";\nEND;";
        }

        private sealed class VariableVisitor : TSqlFragmentVisitor
        {
            private readonly HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            public List<string> Used { get; } = new List<string>();

            public HashSet<string> Declared { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            public override void ExplicitVisit(VariableReference node)
            {
                if (seen.Add(node.Name))
                    Used.Add(node.Name);

                base.ExplicitVisit(node);
            }

            public override void ExplicitVisit(DeclareVariableElement node)
            {
                Declared.Add(node.VariableName.Value);
                base.ExplicitVisit(node);
            }

            public override void ExplicitVisit(DeclareTableVariableBody node)
            {
                if (node.VariableName != null)
                    Declared.Add(node.VariableName.Value);
                base.ExplicitVisit(node);
            }
        }
    }
}
