using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;

namespace Flank.Ssrs
{
    /// <summary>
    /// A parameter's choices (a dropdown): from a stored procedure or a query,
    /// one of whose columns is the value passed and another (or the same) the
    /// label shown, or from values typed in.
    /// </summary>
    public static class ProcedureList
    {
        /// <summary>Choices kept from the preview run, for a workbook to start with.</summary>
        public const int MaxRows = 5000;

        /// <summary>Column names for typed-in values.</summary>
        public const string ValueColumn = "value";
        public const string LabelColumn = "label";

        /// <summary>A list procedure's or query's columns and choices, from a run.</summary>
        public sealed class Description
        {
            public Description(string procedure, IReadOnlyList<string> columns, IReadOnlyList<object?[]> rows)
            {
                Procedure = procedure;
                Columns = columns;
                Rows = rows;
            }

            /// <summary>[schema].[name], or null for a query.</summary>
            public string? Procedure { get; }

            /// <summary>For a query: its text.</summary>
            public string? Query { get; private set; }

            public IReadOnlyList<string> Columns { get; }

            public IReadOnlyList<object?[]> Rows { get; }

            internal static Description ForQuery(string query, IReadOnlyList<string> columns, IReadOnlyList<object?[]> rows) =>
                new Description(null!, columns, rows) { Query = query };
        }

        /// <summary>
        /// Checks the procedure can supply a list (it exists, needs no input
        /// values, returns named columns) and runs it for its columns and choices.
        /// </summary>
        /// <exception cref="InvalidOperationException">With a message for the user if it can't.</exception>
        public static Description Describe(string procedureName, IDbConnection connection)
        {
            int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(procedureName, connection);
            string qualified = QualifiedName(objectId, connection);

            var needsValues = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(objectId, connection)
                .Where(p => !p.IsOutput && !p.HasDefault)
                .Select(p => p.Name)
                .ToList();

            if (needsValues.Count > 0)
            {
                throw new InvalidOperationException(
                    $"{qualified} needs values for {string.Join(", ", needsValues)}, so it can't supply a list. " +
                    "Use a procedure that takes no parameters, or a query that calls it with values.");
            }

            var result = StoredProcedureRunner.Run(
                connection,
                qualified,
                new KeyValuePair<string, object?>[0],
                MaxRows).FirstOrDefault();

            CheckColumns(result, qualified);
            return new Description(qualified, result!.Columns, result.Rows);
        }

        /// <summary>
        /// Checks the query can supply a list (it parses, has no parameters,
        /// returns named columns) and runs it for its columns and choices.
        /// </summary>
        /// <exception cref="InvalidOperationException">With a message for the user if it can't.</exception>
        public static Description DescribeQuery(string query, IDbConnection connection)
        {
            if (string.IsNullOrWhiteSpace(query))
                throw new InvalidOperationException("Write a query that returns the choices, e.g. SELECT id, name FROM dbo.leagues.");

            string? problem = QueryParameters.ParseProblem(query);
            if (problem != null)
                throw new InvalidOperationException("The query doesn't parse. " + problem);

            var parameters = QueryParameters.Find(query);
            if (parameters.Count > 0)
            {
                throw new InvalidOperationException(
                    $"The query uses {string.Join(", ", parameters)}, which nothing would fill in. Use values, or DECLARE them.");
            }

            var result = StoredProcedureRunner.RunQuery(connection, query, MaxRows).FirstOrDefault();

            CheckColumns(result, "The query");
            return Description.ForQuery(query, result!.Columns, result.Rows);
        }

        private static void CheckColumns(ResultSet? result, string what)
        {
            if (result == null || result.Columns.Count == 0 || result.Columns.Any(string.IsNullOrWhiteSpace))
                throw new InvalidOperationException($"{what} must return a result set with named columns.");

            if (result.Columns.Distinct(StringComparer.OrdinalIgnoreCase).Count() != result.Columns.Count)
                throw new InvalidOperationException($"{what} returns two columns with the same name.");
        }

        /// <summary>Sets the list on the settings: the procedure or query, the chosen columns, and their choices.</summary>
        public static void Apply(ParameterSettings settings, Description list, string valueColumn, string labelColumn)
        {
            int value = IndexOf(list, valueColumn);
            int label = IndexOf(list, labelColumn);

            Clear(settings);
            settings.ListSource = list.Query != null ? ListSource.Query : ListSource.Procedure;
            settings.ListProcedure = list.Procedure;
            settings.ListQuery = list.Query;
            settings.ListValueColumn = list.Columns[value];
            settings.ListLabelColumn = list.Columns[label];
            settings.ListRows = list.Rows
                .Select(row => new KeyValuePair<object?, object?>(row[value], row[label]))
                .ToList();
        }

        /// <summary>
        /// Describe and Apply in one, with the first column as the value and the
        /// second (if any) as the label.
        /// </summary>
        public static void Apply(ParameterSettings settings, string procedureName, IDbConnection connection)
        {
            var list = Describe(procedureName, connection);
            Apply(settings, list, list.Columns[0], list.Columns[list.Columns.Count > 1 ? 1 : 0]);
        }

        /// <summary>
        /// Typed-in choices, one per line: the values, and optionally a label
        /// for each (same number of lines). Values must be of the parameter's
        /// type; blank lines are ignored.
        /// </summary>
        /// <param name="labels">Null for values only (each value is its own label).</param>
        /// <param name="problem">Why the lines can't be used, if they can't.</param>
        public static IReadOnlyList<KeyValuePair<object?, object?>>? ParseValues(
            SprocParameter parameter,
            string values,
            string? labels,
            out string? problem)
        {
            var valueLines = Lines(values);
            var labelLines = labels == null ? null : Lines(labels);
            problem = null;

            if (valueLines.Count == 0)
            {
                problem = "Enter at least one value, one per line.";
                return null;
            }

            if (labelLines != null && labelLines.Count == 0)
            {
                problem = $"Type a label for each of the {valueLines.Count} value(s), one per line, in the same order.";
                return null;
            }

            if (labelLines != null && labelLines.Count != valueLines.Count)
            {
                problem = $"There are {valueLines.Count} value(s) but {labelLines.Count} label(s). " +
                          "Give each value a label, on the same line number.";
                return null;
            }

            var rows = new List<KeyValuePair<object?, object?>>();

            for (int i = 0; i < valueLines.Count; i++)
            {
                object? value = ParseTyped(parameter, valueLines[i], out string? why);

                if (why != null)
                {
                    problem = $"Value {i + 1}, \"{valueLines[i]}\": {why}";
                    return null;
                }

                rows.Add(new KeyValuePair<object?, object?>(value, labelLines?[i] ?? valueLines[i]));
            }

            string? duplicate = FirstDuplicate(rows.Select(r => FormatValue(r.Key)));
            if (duplicate != null)
            {
                problem = $"\"{duplicate}\" is in the values twice.";
                return null;
            }

            duplicate = FirstDuplicate(rows.Select(r => (string)r.Value!));
            if (duplicate != null)
            {
                problem = $"\"{duplicate}\" is in the labels twice; people pick by label, so each must be different.";
                return null;
            }

            return rows;
        }

        /// <summary>Sets typed-in choices on the settings (from <see cref="ParseValues"/>).</summary>
        public static void ApplyValues(ParameterSettings settings, IReadOnlyList<KeyValuePair<object?, object?>> rows, bool withLabels)
        {
            Clear(settings);
            settings.ListSource = ListSource.Values;
            settings.ListValueColumn = ValueColumn;
            settings.ListLabelColumn = withLabels ? LabelColumn : ValueColumn;
            settings.ListRows = rows;
        }

        /// <summary>Whether typed-in choices have labels of their own.</summary>
        public static bool HasLabels(ParameterSettings settings) =>
            settings.ListSource == ListSource.Values && settings.ListLabelColumn != settings.ListValueColumn;

        /// <summary>A value as it's typed in: invariant numbers, 1/0 for bit, yyyy-MM-dd dates.</summary>
        public static string FormatValue(object? value)
        {
            switch (value)
            {
                case null:
                    return "";
                case bool flag:
                    return flag ? "1" : "0";
                case DateTime date:
                    return date.TimeOfDay == TimeSpan.Zero
                        ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                case IFormattable formattable:
                    return formattable.ToString(null, CultureInfo.InvariantCulture);
                default:
                    return value.ToString() ?? "";
            }
        }

        private static List<string> Lines(string text) =>
            (text ?? "")
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
                .Select(line => line.Trim())
                .Where(line => line.Length > 0)
                .ToList();

        private static string? FirstDuplicate(IEnumerable<string> items)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            return items.FirstOrDefault(item => !seen.Add(item));
        }

        /// <summary>The value as the parameter's type, or why it isn't one.</summary>
        private static object? ParseTyped(SprocParameter parameter, string text, out string? why)
        {
            string type = (parameter.SqlType ?? "").ToLowerInvariant();
            why = null;

            // Typed-in values read the same whatever the Windows settings: 1.5, 2024-01-31.
            object? value = ParseInvariant(type, text) ?? StoredProcedureRunner.ParseValue(type, text);

            switch (type)
            {
                case "tinyint":
                case "smallint":
                case "int":
                case "bigint":
                    if (!(value is long whole))
                    {
                        why = "not a whole number.";
                        return null;
                    }

                    var (min, max) = type == "tinyint" ? (0L, (long)byte.MaxValue)
                        : type == "smallint" ? ((long)short.MinValue, (long)short.MaxValue)
                        : type == "int" ? ((long)int.MinValue, (long)int.MaxValue)
                        : (long.MinValue, long.MaxValue);

                    if (whole < min || whole > max)
                        why = $"out of range for {type} ({min} to {max}).";

                    return whole;

                case "decimal":
                case "numeric":
                case "money":
                case "smallmoney":
                case "float":
                case "real":
                    if (!(value is decimal) && !(value is double))
                        why = "not a number.";
                    return value;

                case "bit":
                    if (!(value is bool))
                        why = "not 1, 0, true or false.";
                    return value;

                case "date":
                case "datetime":
                case "datetime2":
                case "smalldatetime":
                    if (!(value is DateTime))
                        why = "not a date.";
                    return value;

                case "uniqueidentifier":
                    if (!Guid.TryParse(text, out _))
                        why = "not a GUID.";
                    return text;

                case "char":
                case "varchar":
                case "nchar":
                case "nvarchar":
                    int length = parameter.MaxLength < 0 ? int.MaxValue
                        : type.StartsWith("n", StringComparison.Ordinal) ? parameter.MaxLength / 2
                        : parameter.MaxLength;

                    if (text.Length > length)
                        why = $"longer than {ProcedureCall.TypeDeclaration(parameter)} allows ({length} characters).";

                    return text;

                default:
                    return text;
            }
        }

        private static object? ParseInvariant(string type, string text)
        {
            var invariant = CultureInfo.InvariantCulture;

            switch (type)
            {
                case "decimal":
                case "numeric":
                case "money":
                case "smallmoney":
                    return decimal.TryParse(text, NumberStyles.Number, invariant, out decimal number) ? number : (object?)null;

                case "float":
                case "real":
                    return double.TryParse(text, NumberStyles.Float, invariant, out double real) ? real : (object?)null;

                case "date":
                case "datetime":
                case "datetime2":
                case "smalldatetime":
                    return DateTime.TryParseExact(
                        text,
                        new[] { "yyyy-MM-dd", "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss" },
                        invariant,
                        DateTimeStyles.None,
                        out DateTime date) ? date : (object?)null;

                default:
                    return null;
            }
        }

        /// <summary>
        /// The database's own procedures that take no input parameters, as
        /// [schema].[name]: the ones that can supply a list as they are.
        /// (Others can through a query that calls them with values.)
        /// </summary>
        public static IReadOnlyList<string> Candidates(IDbConnection connection)
        {
            var names = new List<string>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT QUOTENAME(SCHEMA_NAME(p.schema_id)) + '.' + QUOTENAME(p.name)
FROM sys.procedures AS p
WHERE p.is_ms_shipped = 0
  AND NOT EXISTS (SELECT 1 FROM sys.parameters AS pa WHERE pa.object_id = p.object_id AND pa.is_output = 0)
ORDER BY SCHEMA_NAME(p.schema_id), p.name;";

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                        names.Add(reader.GetString(0));
                }
            }

            return names;
        }

        /// <summary>Removes the list from the settings.</summary>
        public static void Clear(ParameterSettings settings)
        {
            settings.ListSource = ListSource.Procedure;
            settings.ListProcedure = null;
            settings.ListQuery = null;
            settings.ListValueColumn = null;
            settings.ListLabelColumn = null;
            settings.ListRows = new KeyValuePair<object?, object?>[0];
        }

        private static int IndexOf(Description list, string column)
        {
            for (int i = 0; i < list.Columns.Count; i++)
            {
                if (string.Equals(list.Columns[i], column, StringComparison.OrdinalIgnoreCase))
                    return i;
            }

            throw new InvalidOperationException($"{list.Procedure ?? "The query"} has no column {column}.");
        }

        private static string QualifiedName(int objectId, IDbConnection connection)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandText =
                    "SELECT QUOTENAME(OBJECT_SCHEMA_NAME(@ObjectId)) + '.' + QUOTENAME(OBJECT_NAME(@ObjectId));";

                var parameter = command.CreateParameter();
                parameter.ParameterName = "@ObjectId";
                parameter.Value = objectId;
                command.Parameters.Add(parameter);

                return (string)command.ExecuteScalar();
            }
        }
    }
}
