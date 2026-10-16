using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using System;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// The SQL login (SQL Server authentication) an SSMS connection uses, to
    /// pre-fill the login a Power BI report reads the database with. Null for
    /// Windows and Microsoft Entra connections, which have no password to reuse.
    /// </summary>
    internal sealed class SsmsSqlLogin
    {
        // UIConnectionInfo.AuthenticationType for SQL Server authentication.
        private const int SqlAuthentication = 1;

        private SsmsSqlLogin(string username, string password)
        {
            Username = username;
            Password = password;
        }

        public string Username { get; }

        public string Password { get; }

        /// <summary>The active query window's login.</summary>
        public static SsmsSqlLogin? FromActiveQueryWindow()
        {
            try
            {
                var info = ServiceCache.ScriptFactory?.CurrentlyActiveWndConnectionInfo?.UIConnectionInfo;

                if (info == null || info.AuthenticationType != SqlAuthentication)
                    return null;

                return Create(info.UserName, info.Password);
            }
            catch (Exception ex)
            {
                // Only a convenience: the user can type the login instead.
                System.Diagnostics.Debug.WriteLine($"Flank: could not read the query window's login: {ex.Message}");
                return null;
            }
        }

        /// <summary>The active query window's server, as SSMS connected to it.</summary>
        public static string? ActiveQueryWindowServer()
        {
            try
            {
                return ServiceCache.ScriptFactory?.CurrentlyActiveWndConnectionInfo?.UIConnectionInfo?.ServerName;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>An Object Explorer connection's login.</summary>
        public static SsmsSqlLogin? FromConnectionInfo(SqlConnectionInfo info)
        {
            if (info.UseIntegratedSecurity)
                return null;

            if (info.Authentication != SqlConnectionInfo.AuthenticationMethod.NotSpecified &&
                info.Authentication != SqlConnectionInfo.AuthenticationMethod.SqlPassword)
            {
                return null;
            }

            return Create(info.UserName, info.Password);
        }

        private static SsmsSqlLogin? Create(string? username, string? password) =>
            string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password)
                ? null
                : new SsmsSqlLogin(username!, password!);
    }
}
