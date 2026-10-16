using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Flank.Ssrs
{
    /// <summary>
    /// A stored procedure that supplies a parameter's choices (a dropdown):
    /// its first column is the value, its second (if any) the label.
    /// </summary>
    public static class ProcedureList
    {
        /// <summary>Choices kept from the preview run, for a workbook to start with.</summary>
        public const int MaxRows = 5000;

        /// <summary>
        /// Checks the procedure can supply a list (it exists, needs no input
        /// values, returns at least one column) and sets the list on the
        /// settings, with its choices from a run now.
        /// </summary>
        /// <exception cref="InvalidOperationException">With a message for the user if it can't.</exception>
        public static void Apply(ParameterSettings settings, string procedureName, IDbConnection connection)
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
            {
                throw new InvalidOperationException(
                    $"{qualified} must return a result set with named columns: the value, then optionally a label.");
            }

            string valueColumn = result.Columns[0];
            string labelColumn = result.Columns.Count > 1 ? result.Columns[1] : valueColumn;

            settings.ListProcedure = qualified;
            settings.ListValueColumn = valueColumn;
            settings.ListLabelColumn = labelColumn;
            settings.ListRows = result.Rows
                .Select(row => new KeyValuePair<object?, object?>(row[0], result.Columns.Count > 1 ? row[1] : row[0]))
                .ToList();
        }

        /// <summary>Removes the list from the settings.</summary>
        public static void Clear(ParameterSettings settings)
        {
            settings.ListProcedure = null;
            settings.ListValueColumn = null;
            settings.ListLabelColumn = null;
            settings.ListRows = new KeyValuePair<object?, object?>[0];
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
