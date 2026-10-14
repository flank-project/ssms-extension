using System;
using System.Collections.Generic;
using System.Data;

namespace Flank.Ssrs
{
    /// <summary>Facts about the server and query that decide how to grant access.</summary>
    public static class AccessDiscovery
    {
        /// <summary>
        /// Azure SQL Database (engine edition 5) has no server logins for SQL
        /// authentication; users are created in the database.
        /// </summary>
        public static bool IsAzureSqlDatabase(IDbConnection connection)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandText = "SELECT CAST(SERVERPROPERTY('EngineEdition') AS int)";
                return Convert.ToInt32(command.ExecuteScalar()) == 5;
            }
        }

        /// <summary>
        /// The tables and views in this database that a query's result columns
        /// come from, as [schema].[name], without running the query. Tables used
        /// only in WHERE or JOIN conditions aren't listed.
        /// </summary>
        public static IReadOnlyList<string> QueryTables(IDbConnection connection, string sql)
        {
            var tables = new List<string>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
EXEC sp_describe_first_result_set
    @tsql = @Sql,
    @params = NULL,
    @browse_information_mode = 1;";

                var parameter = command.CreateParameter();
                parameter.ParameterName = "@Sql";
                parameter.Value = sql;
                command.Parameters.Add(parameter);

                using (var reader = command.ExecuteReader())
                {
                    int database = reader.GetOrdinal("source_database");
                    int schema = reader.GetOrdinal("source_schema");
                    int table = reader.GetOrdinal("source_table");
                    int hidden = reader.GetOrdinal("is_hidden");

                    while (reader.Read())
                    {
                        if (reader.IsDBNull(table) || reader.IsDBNull(schema))
                            continue;

                        // Hidden columns are browse-mode keys, not part of the result.
                        if (!reader.IsDBNull(hidden) && reader.GetBoolean(hidden))
                            continue;

                        // Grants only apply in this database.
                        if (!reader.IsDBNull(database) &&
                            !string.Equals(reader.GetString(database), connection.Database, StringComparison.OrdinalIgnoreCase))
                            continue;

                        string name = AccessScript.QuoteName(reader.GetString(schema)) + "." +
                                      AccessScript.QuoteName(reader.GetString(table));

                        if (!tables.Contains(name))
                            tables.Add(name);
                    }
                }
            }

            return tables;
        }
    }
}
