using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Flank.Ssrs
{
    /// <summary>How the end user signs in to the database.</summary>
    public enum AccessSignIn
    {
        /// <summary>A SQL Server username and password.</summary>
        SqlLogin,

        /// <summary>A Microsoft Entra account (Azure SQL).</summary>
        Entra,

        /// <summary>A Windows / Active Directory account (SQL Server).</summary>
        Windows,
    }

    /// <summary>What the end user is allowed to do.</summary>
    public enum AccessScope
    {
        /// <summary>EXECUTE on one stored procedure.</summary>
        Procedure,

        /// <summary>SELECT on specific tables and views.</summary>
        Objects,
    }

    /// <summary>A request to give someone access, and the T-SQL that does it.</summary>
    public sealed class AccessRequest
    {
        public AccessSignIn SignIn { get; set; }

        /// <summary>Login/user name: a SQL login name, an Entra UPN, or DOMAIN\user.</summary>
        public string Name { get; set; } = "";

        /// <summary>The user already exists: only grant access.</summary>
        public bool ExistingUser { get; set; }

        /// <summary>For <see cref="AccessSignIn.SqlLogin"/>.</summary>
        public string Password { get; set; } = "";

        /// <summary>The database to give access to.</summary>
        public string Database { get; set; } = "";

        /// <summary>
        /// Azure SQL Database (not Managed Instance): users are created in the
        /// database itself (contained users), with no server login.
        /// </summary>
        public bool IsAzureSqlDatabase { get; set; }

        public AccessScope Scope { get; set; }

        /// <summary>For <see cref="AccessScope.Procedure"/>: [schema].[name].</summary>
        public string Procedure { get; set; } = "";

        /// <summary>
        /// Create <see cref="Procedure"/> from this query first (turning a query
        /// into a procedure). Null when the procedure already exists.
        /// </summary>
        public string? ProcedureBody { get; set; }

        /// <summary>For <see cref="AccessScope.Objects"/>: [schema].[name] of each table or view.</summary>
        public IReadOnlyList<string> Objects { get; set; } = Array.Empty<string>();
    }

    public static class AccessScript
    {
        /// <summary>
        /// The statements that create the user (and login, where needed) and
        /// grant the access. Each statement is idempotent where SQL Server
        /// allows: existing logins/users are left alone, grants are repeatable.
        /// </summary>
        public static IReadOnlyList<string> Build(AccessRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Name))
                throw new ArgumentException("A name is required.");

            string user = QuoteName(request.Name);
            string userLiteral = QuoteString(request.Name);
            var statements = new List<string>();

            if (request.Scope == AccessScope.Procedure && request.ProcedureBody != null)
            {
                // Must be the only statement in its batch.
                statements.Add(
                    $"CREATE PROCEDURE {request.Procedure}\nAS\nBEGIN\n    SET NOCOUNT ON;\n\n" +
                    request.ProcedureBody.Trim().TrimEnd(';') + ";\nEND;");
            }

            switch (request.ExistingUser ? (AccessSignIn?)null : request.SignIn)
            {
                case AccessSignIn.SqlLogin:
                    if (string.IsNullOrEmpty(request.Password))
                        throw new ArgumentException("A password is required.");

                    if (request.IsAzureSqlDatabase)
                    {
                        // Contained database user: signs in to this database only.
                        statements.Add(
                            $"IF DATABASE_PRINCIPAL_ID({userLiteral}) IS NULL\n" +
                            $"    CREATE USER {user} WITH PASSWORD = {QuoteString(request.Password)};");
                    }
                    else
                    {
                        statements.Add(
                            $"IF SUSER_ID({userLiteral}) IS NULL\n" +
                            $"    CREATE LOGIN {user} WITH PASSWORD = {QuoteString(request.Password)}, DEFAULT_DATABASE = {QuoteName(request.Database)};");
                        statements.Add(
                            $"IF DATABASE_PRINCIPAL_ID({userLiteral}) IS NULL\n" +
                            $"    CREATE USER {user} FOR LOGIN {user};");
                    }
                    break;

                case AccessSignIn.Entra:
                    statements.Add(
                        $"IF DATABASE_PRINCIPAL_ID({userLiteral}) IS NULL\n" +
                        $"    CREATE USER {user} FROM EXTERNAL PROVIDER;");
                    break;

                case AccessSignIn.Windows:
                    statements.Add(
                        $"IF SUSER_ID({userLiteral}) IS NULL\n" +
                        $"    CREATE LOGIN {user} FROM WINDOWS WITH DEFAULT_DATABASE = {QuoteName(request.Database)};");
                    statements.Add(
                        $"IF DATABASE_PRINCIPAL_ID({userLiteral}) IS NULL\n" +
                        $"    CREATE USER {user} FOR LOGIN {user};");
                    break;
            }

            switch (request.Scope)
            {
                case AccessScope.Procedure:
                    statements.Add($"GRANT EXECUTE ON OBJECT::{request.Procedure} TO {user};");
                    break;

                case AccessScope.Objects:
                    if (request.Objects.Count == 0)
                        throw new ArgumentException("No tables or views to grant.");
                    foreach (var name in request.Objects)
                        statements.Add($"GRANT SELECT ON OBJECT::{name} TO {user};");
                    break;

            }

            return statements;
        }

        /// <summary>The whole script as text, for review or copying.</summary>
        public static string ToText(AccessRequest request)
        {
            var text = new StringBuilder();

            if (!request.IsAzureSqlDatabase)
                text.AppendLine($"USE {QuoteName(request.Database)};").AppendLine();

            // Statements break lines with \n; text boxes and editors want CRLF.
            // GO separates batches, so the script can be pasted and run as is
            // (CREATE PROCEDURE must be alone in its batch).
            foreach (var statement in Build(request))
                text.AppendLine(statement.Replace("\n", Environment.NewLine)).AppendLine("GO").AppendLine();

            return text.ToString().TrimEnd() + Environment.NewLine;
        }

        /// <summary>Runs the statements on a connection to the target database.</summary>
        public static void Run(IDbConnection connection, AccessRequest request)
        {
            foreach (var statement in Build(request))
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = statement;
                    command.ExecuteNonQuery();
                }
            }
        }

        /// <summary>A random password that meets SQL Server's complexity policy.</summary>
        public static string GeneratePassword(int length = 20)
        {
            const string upper = "ABCDEFGHJKLMNPQRSTUVWXYZ";
            const string lower = "abcdefghijkmnopqrstuvwxyz";
            const string digits = "23456789";
            const string symbols = "!#$%*+-=?@^_";
            string all = upper + lower + digits + symbols;

            var chars = new List<char>
            {
                Pick(upper), Pick(lower), Pick(digits), Pick(symbols),
            };

            while (chars.Count < length)
                chars.Add(Pick(all));

            // Shuffle so the required classes aren't always first.
            for (int i = chars.Count - 1; i > 0; i--)
            {
                int j = RandomIndex(i + 1);
                (chars[i], chars[j]) = (chars[j], chars[i]);
            }

            return new string(chars.ToArray());
        }

        /// <summary>[name] with ] doubled.</summary>
        public static string QuoteName(string name) => "[" + name.Replace("]", "]]") + "]";

        /// <summary>N'text' with ' doubled.</summary>
        public static string QuoteString(string text) => "N'" + text.Replace("'", "''") + "'";

        private static char Pick(string set) => set[RandomIndex(set.Length)];

        private static int RandomIndex(int count)
        {
            var bytes = new byte[4];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }
            return (int)(BitConverter.ToUInt32(bytes, 0) % (uint)count);
        }
    }
}
