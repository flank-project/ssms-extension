using System;
using System.Collections.Generic;
using System.Data;

namespace Flank.Ssrs
{
    public class SqlMetadataDiscovery
    {
        public static IReadOnlyList<ReportColumn> DiscoverQueryColumns(
            string sql,
            IDbConnection connection) =>
            DiscoverQueryColumns(sql, connection, new SprocParameter[0]);

        /// <param name="parameters">The query's undeclared @names, with their types (passed as NULL).</param>
        public static IReadOnlyList<ReportColumn> DiscoverQueryColumns(
            string sql,
            IDbConnection connection,
            IReadOnlyList<SprocParameter> parameters)
        {
            var columns = new List<ReportColumn>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = sql;
                StoredProcedureRunner.AddQueryParameters(command, parameters, null);

                using (var reader = command.ExecuteReader(
                    System.Data.CommandBehavior.SchemaOnly))
                {
                    // Null when the query returns no result set.
                    var schema = reader.GetSchemaTable();

                    foreach (System.Data.DataRow row in schema?.Rows ?? (System.Collections.IEnumerable)Array.Empty<DataRow>())
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
    p.is_output,
    CASE WHEN t.is_table_type = 1
        THEN QUOTENAME(SCHEMA_NAME(t.schema_id)) + N'.' + QUOTENAME(t.name)
    END AS table_type
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
                            IsOutput = reader.GetBoolean(5),
                            TableType = reader.IsDBNull(6) ? null : reader.GetString(6),
                        });
                    }
                }
            }

            foreach (var p in parameters)
            {
                if (p.IsTable)
                    p.Columns = DiscoverTableTypeColumns(p.TableType!, connection);
            }

            if (parameters.Count > 0)
            {
                // T-SQL defaults aren't in sys.parameters; read them from the definition.
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT OBJECT_DEFINITION(@ObjectId);";

                    var parameter = command.CreateParameter();
                    parameter.ParameterName = "@ObjectId";
                    parameter.Value = objectId;
                    command.Parameters.Add(parameter);

                    var defaults = ProcedureDefaults.Parse(command.ExecuteScalar() as string);

                    foreach (var p in parameters)
                    {
                        if (defaults.TryGetValue(p.Name, out var declared))
                        {
                            p.HasDefault = true;
                            p.DefaultValue = declared.Value;
                            p.DefaultText = declared.Text;
                        }
                    }
                }
            }

            return parameters;
        }

        /// <summary>
        /// The columns of a table type that take values, described like
        /// parameters. Identity, computed and rowversion columns are left out:
        /// SQL Server fills them in. Alias types are reported as their base type.
        /// </summary>
        public static IReadOnlyList<SprocParameter> DiscoverTableTypeColumns(
            string tableType,
            IDbConnection connection)
        {
            var columns = new List<SprocParameter>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT
    c.name,
    TYPE_NAME(c.system_type_id) AS type_name,
    c.max_length,
    c.precision,
    c.scale,
    c.is_nullable
FROM sys.table_types tt
JOIN sys.columns c
    ON c.object_id = tt.type_table_object_id
WHERE tt.user_type_id = TYPE_ID(@TableType)
  AND c.is_identity = 0
  AND c.is_computed = 0
  AND TYPE_NAME(c.system_type_id) <> N'timestamp'
ORDER BY c.column_id;";

                var parameter = command.CreateParameter();
                parameter.ParameterName = "@TableType";
                parameter.Value = tableType;
                command.Parameters.Add(parameter);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columns.Add(new SprocParameter
                        {
                            Name = reader.GetString(0),
                            SqlType = reader.GetString(1),
                            MaxLength = reader.GetInt16(2),
                            Precision = reader.GetByte(3),
                            Scale = reader.GetByte(4),
                            IsNullable = reader.GetBoolean(5),
                        });
                    }
                }
            }

            return columns;
        }

        /// <summary>
        /// The database's compatibility level. Flank passes table-valued
        /// parameters as JSON, read with OPENJSON, which needs 130 (SQL Server 2016) or later.
        /// </summary>
        public static int GetCompatibilityLevel(IDbConnection connection)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandText = "SELECT compatibility_level FROM sys.databases WHERE name = DB_NAME();";
                return Convert.ToInt32(command.ExecuteScalar());
            }
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