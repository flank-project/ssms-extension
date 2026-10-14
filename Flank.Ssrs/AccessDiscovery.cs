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
        /// Database users who sign in the given way, for granting access to
        /// someone who already has a user. Leaves out built-in principals.
        /// </summary>
        public static IReadOnlyList<string> ListUsers(IDbConnection connection, AccessSignIn signIn)
        {
            // S = SQL user, E/X = Entra user/group, U/G = Windows user/group.
            string types = signIn == AccessSignIn.Entra ? "'E', 'X'"
                         : signIn == AccessSignIn.Windows ? "'U', 'G'"
                         : "'S'";

            var users = new List<string>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = $@"
SELECT name
FROM sys.database_principals
WHERE type IN ({types})
  AND principal_id > 4
  AND name NOT IN ('dbo', 'guest', 'INFORMATION_SCHEMA', 'sys')
  AND is_fixed_role = 0
ORDER BY name;";

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                        users.Add(reader.GetString(0));
                }
            }

            return users;
        }
    }
}
