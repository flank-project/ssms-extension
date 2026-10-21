using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Flank.Ssrs
{
    /// <summary>
    /// A stored procedure that supplies a parameter's choices (a dropdown):
    /// one of its columns is the value passed, another (or the same) the label shown.
    /// </summary>
    public static class ProcedureList
    {
        /// <summary>Choices kept from the preview run, for a workbook to start with.</summary>
        public const int MaxRows = 5000;

        /// <summary>A list procedure's columns and choices, from a run.</summary>
        public sealed class Description
        {
            public Description(string procedure, IReadOnlyList<string> columns, IReadOnlyList<object?[]> rows)
            {
                Procedure = procedure;
                Columns = columns;
                Rows = rows;
            }

            /// <summary>[schema].[name].</summary>
            public string Procedure { get; }

            public IReadOnlyList<string> Columns { get; }

            public IReadOnlyList<object?[]> Rows { get; }
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
                    "Use a procedure that takes no parameters (or only ones with defaults).");
            }

            var result = StoredProcedureRunner.Run(
                connection,
                qualified,
                new KeyValuePair<string, object?>[0],
                MaxRows).FirstOrDefault();

            if (result == null || result.Columns.Count == 0 || result.Columns.Any(string.IsNullOrWhiteSpace))
                throw new InvalidOperationException($"{qualified} must return a result set with named columns.");

            if (result.Columns.Distinct(StringComparer.OrdinalIgnoreCase).Count() != result.Columns.Count)
                throw new InvalidOperationException($"{qualified} returns two columns with the same name.");

            return new Description(qualified, result.Columns, result.Rows);
        }

        /// <summary>Sets the list on the settings: the procedure, the chosen columns, and their choices.</summary>
        public static void Apply(ParameterSettings settings, Description list, string valueColumn, string labelColumn)
        {
            int value = IndexOf(list, valueColumn);
            int label = IndexOf(list, labelColumn);

            settings.ListProcedure = list.Procedure;
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
        /// The database's own procedures that take no input parameters, as
        /// [schema].[name]: the ones that can obviously supply a list. (Ones
        /// whose parameters all have defaults can too, but telling needs their
        /// definitions parsed, so they're typed in rather than listed.)
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
            settings.ListProcedure = null;
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

            throw new InvalidOperationException($"{list.Procedure} has no column {column}.");
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
