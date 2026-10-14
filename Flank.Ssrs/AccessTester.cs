using System;
using System.Collections.Generic;
using System.Data;

namespace Flank.Ssrs
{
    /// <summary>One permission to check, e.g. EXECUTE on [dbo].[proc].</summary>
    public sealed class PermissionCheck
    {
        public PermissionCheck(string objectName, string permission)
        {
            ObjectName = objectName;
            Permission = permission;
        }

        /// <summary>[schema].[name].</summary>
        public string ObjectName { get; }

        /// <summary>EXECUTE or SELECT.</summary>
        public string Permission { get; }

        public bool Granted { get; internal set; }
    }

    /// <summary>
    /// Checks what a database user can do by impersonating them. Asks SQL
    /// Server (HAS_PERMS_BY_NAME) rather than running anything as them.
    /// </summary>
    public static class AccessTester
    {
        /// <summary>
        /// Fills in <see cref="PermissionCheck.Granted"/> for each check, as
        /// <paramref name="userName"/>. Needs IMPERSONATE on that user (e.g.
        /// db_owner). The connection is always reverted to its own identity,
        /// so a borrowed connection is safe to use.
        /// </summary>
        public static void Check(IDbConnection connection, string userName, IReadOnlyList<PermissionCheck> checks)
        {
            Execute(connection, $"EXECUTE AS USER = {AccessScript.QuoteString(userName)};");

            try
            {
                foreach (var check in checks)
                {
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = "SELECT HAS_PERMS_BY_NAME(@Object, 'OBJECT', @Permission);";
                        AddParameter(command, "@Object", check.ObjectName);
                        AddParameter(command, "@Permission", check.Permission);

                        // NULL means the user can't even see the object.
                        object result = command.ExecuteScalar();
                        check.Granted = result != null && result != DBNull.Value && Convert.ToInt32(result) == 1;
                    }
                }
            }
            finally
            {
                Execute(connection, "REVERT;");
            }
        }

        /// <summary>Whether a user with this name exists in the database.</summary>
        public static bool UserExists(IDbConnection connection, string userName)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandText = "SELECT DATABASE_PRINCIPAL_ID(@Name);";
                AddParameter(command, "@Name", userName);

                object result = command.ExecuteScalar();
                return result != null && result != DBNull.Value;
            }
        }

        private static void Execute(IDbConnection connection, string sql)
        {
            using (var command = connection.CreateCommand())
            {
                command.CommandText = sql;
                command.ExecuteNonQuery();
            }
        }

        private static void AddParameter(IDbCommand command, string name, string value)
        {
            var parameter = command.CreateParameter();
            parameter.ParameterName = name;
            parameter.Value = value;
            command.Parameters.Add(parameter);
        }
    }
}
