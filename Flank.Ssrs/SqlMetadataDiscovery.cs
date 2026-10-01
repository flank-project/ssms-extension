using System;
using System.Collections.Generic;
using System.Data;

namespace Flank.Ssrs
{
    public class SqlMetadataDiscovery
    {
        public static IReadOnlyList<ReportColumn> DiscoverQueryColumns(
            string sql,
            IDbConnection connection)
        {
            var columns = new List<ReportColumn>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = sql;

                using (var reader = command.ExecuteReader(
                    System.Data.CommandBehavior.SchemaOnly))
                {
                    var schema = reader.GetSchemaTable();

                    foreach (System.Data.DataRow row in schema.Rows)
                    {
                        columns.Add(new ReportColumn
                        {
                            Name = (string)row["ColumnName"]
                        });
                    }
                }
            }

            if (columns.Count == 0)
            {
                throw new InvalidOperationException(
                    "The query does not return any columns.");
            }

            return columns;
        }

        public static int GetStoredProcedureObjectId(
            string procedureName,
            IDbConnection connection)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT OBJECT_ID(@ProcedureName, 'P')";

                var parameter = command.CreateParameter();
                parameter.ParameterName = "@ProcedureName";
                parameter.Value = procedureName;
                command.Parameters.Add(parameter);


                var result = command.ExecuteScalar();

                if (result == null ||
                    result == DBNull.Value)
                {
                    throw new InvalidOperationException(
                        $"Stored procedure '{procedureName}' not found.");
                }

                return (int)result;
            }
        }

        public static IReadOnlyList<SprocParameter> DiscoverStoredProcedureParameters(
            int objectId,
            IDbConnection connection)
        {
            var parameters = new List<SprocParameter>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT
    p.name,
    t.name AS type_name,
    p.max_length,
    p.precision,
    p.scale,
    p.is_output
FROM sys.parameters p
JOIN sys.types t
    ON p.user_type_id = t.user_type_id
WHERE p.object_id = @ObjectId
ORDER BY p.parameter_id;";

                var parameter = command.CreateParameter();
                parameter.ParameterName = "@ObjectId";
                parameter.Value = objectId;
                command.Parameters.Add(parameter);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        parameters.Add(new SprocParameter
                        {
                            Name = reader.GetString(0),
                            SqlType = reader.GetString(1),
                            MaxLength = reader.GetInt16(2),
                            Precision = reader.GetByte(3),
                            Scale = reader.GetByte(4),
                            IsOutput = reader.GetBoolean(5)
                        });
                    }
                }
            }

            return parameters;
        }

        public static IReadOnlyList<ReportColumn> DiscoverStoredProcedureColumns(
            int objectId,
            string procedureName,
            IDbConnection connection)
        {
            var columns = new List<ReportColumn>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT
    name,
    system_type_name,
    is_nullable,
    column_ordinal
FROM sys.dm_exec_describe_first_result_set_for_object(
    @ObjectId,
    NULL
)
WHERE is_hidden = 0
ORDER BY column_ordinal;";

                var parameter = command.CreateParameter();
                parameter.ParameterName = "@ObjectId";
                parameter.Value = objectId;
                command.Parameters.Add(parameter);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columns.Add(new ReportColumn
                        {
                            Name = reader.IsDBNull(0)
                                ? null
                                : reader.GetString(0)
                        });
                    }
                }
            }

            if (columns.Count == 0)
            {
                throw new InvalidOperationException(
                    $"Stored procedure '{procedureName}' " +
                    "does not return a result set.");
            }

            return columns;
        }
    }
}