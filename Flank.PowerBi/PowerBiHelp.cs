using Microsoft.Identity.Client;
using System;
using System.Text.RegularExpressions;

namespace Flank.PowerBi
{
    /// <summary>The steps of publishing to Power BI, for saying which one failed.</summary>
    public enum PowerBiStep
    {
        SignIn,
        Workspaces,
        Connections,
        Upload,
        Connect,
    }

    /// <summary>
    /// What failed, why, and the setup that usually fixes it. Setup steps are
    /// dated: Microsoft moves its admin menus.
    /// </summary>
    public static class PowerBiHelp
    {
        public const string AsOf = "As of October 21, 2026";

        private const string Permissions =
            "Power BI: Report.ReadWrite.All, Dataset.ReadWrite.All and Workspace.Read.All; " +
            "Fabric: Connection.Read.All";

        /// <summary>A headline for the step, the error itself, and what to check.</summary>
        public static string Explain(PowerBiStep step, Exception error)
        {
            string headline = Headline(step);
            string message = Message(error);
            string? fix = Fix(step, error);

            return headline + "\n\n" + message + (fix == null ? "" : "\n\n" + AsOf + ": " + fix);
        }

        /// <summary>What Flank was doing, as a sentence.</summary>
        public static string Headline(PowerBiStep step)
        {
            switch (step)
            {
                case PowerBiStep.SignIn: return "Flank couldn't sign you in to Power BI.";
                case PowerBiStep.Workspaces: return "Flank couldn't load your Power BI workspaces.";
                case PowerBiStep.Connections: return "Flank couldn't list your Power BI Gateway / Cloud Connections.";
                case PowerBiStep.Upload: return "Power BI didn't accept the report.";
                default: return "Flank uploaded the report but couldn't connect it to the database.";
            }
        }

        /// <summary>When the user has no workspace Flank can publish to.</summary>
        public static string NoWorkspaces =>
            "You aren't a Contributor, Member or Admin of any Power BI workspace, so there's nowhere to publish the report.\n\n" +
            AsOf + ": ask a workspace admin to add you. In Power BI, they open the workspace, choose Manage access, " +
            "and add you as a Contributor (or higher). The workspace needs a Pro or Premium Per User licence for you, " +
            "or Premium / Fabric capacity.";

        /// <summary>When Flank may sign in but not list connections (that permission wasn't approved).</summary>
        public static string ConnectionsNotAllowed =>
            Headline(PowerBiStep.Connections) + " It hasn't been allowed to read them.\n\n" +
            AsOf + ": if signing in again doesn't offer to allow it, your organization needs an admin to approve it. " +
            AdminConsent;

        private static string AdminConsent =>
            "An admin (Cloud Application Administrator or Global Administrator) opens the Microsoft Entra admin center, " +
            "goes to Enterprise applications → Admin consent requests (or Enterprise applications → Flank → Permissions), " +
            "and grants Flank's permissions (" + Permissions + ").";

        private static string Message(Exception error)
        {
            // MSAL messages carry a trace and correlation id after the first line.
            string message = error.Message.Trim();
            int end = message.IndexOf("\nTrace ID", StringComparison.OrdinalIgnoreCase);
            return end > 0 ? message.Substring(0, end).Trim() : message;
        }

        private static string? Fix(PowerBiStep step, Exception error)
        {
            string text = error.ToString();
            string? aadsts = Regex.Match(text, @"AADSTS\d+").Value is var code && code.Length > 0 ? code : null;

            if (error is MsalException || aadsts != null)
            {
                switch (aadsts)
                {
                    case "AADSTS65001":
                    case "AADSTS90094":
                    case "AADSTS90095":
                        return "your organization requires an admin to approve Flank. " + AdminConsent;

                    case "AADSTS65004":
                        return "the permission request was declined. Sign in again and accept it, or ask an admin to approve Flank. " + AdminConsent;

                    case "AADSTS50020":
                    case "AADSTS700016":
                    case "AADSTS90072":
                    case "AADSTS50194":
                        return "this version of Flank can't sign in to your organization's Microsoft Entra tenant. " +
                               "Contact the Flank developer (Make Self-Serve → Feedback...).";

                    case "AADSTS53003":
                        return "a Conditional Access policy in your organization blocked the sign-in. " +
                               "Ask your Microsoft Entra admin to allow Flank (Microsoft Entra admin center → Protection → Conditional Access).";

                    case "AADSTS50076":
                    case "AADSTS50079":
                        return "your organization requires multi-factor authentication. Try again with Sign in (the browser), which can ask for it.";
                }

                return step == PowerBiStep.SignIn
                    ? "if this keeps happening, try Sign in with a code instead of the browser, or ask your admin whether apps may sign in to Power BI."
                    : null;
            }

            bool forbidden = text.Contains("401") || text.Contains("403") ||
                             text.IndexOf("Unauthorized", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             text.IndexOf("Forbidden", StringComparison.OrdinalIgnoreCase) >= 0;

            switch (step)
            {
                case PowerBiStep.Workspaces:
                    return forbidden
                        ? "you need a Power BI licence (Pro or Premium Per User), or a workspace on Premium / Fabric capacity, " +
                          "and Flank's permissions approved. " + AdminConsent
                        : null;

                case PowerBiStep.Connections:
                    return "Flank needs the Fabric Connection.Read.All permission. " + AdminConsent;

                case PowerBiStep.Upload:
                    return "you need the Contributor, Member or Admin role on the workspace (in Power BI: the workspace → Manage access), " +
                           "and the workspace must allow paginated reports: a Pro or Premium Per User licence for you, " +
                           "or Premium / Fabric capacity for the workspace.";

                case PowerBiStep.Connect:
                    return "you must be a user of the Gateway / Cloud Connection. In Power BI: Settings (gear) → Manage connections and gateways " +
                           "→ the connection → Manage users → add yourself as a User. For Azure SQL, the server's firewall must let Power BI in " +
                           "(Azure Portal → the SQL server → Networking → Allow Azure services and resources to access this server). " +
                           "For a gateway connection, the gateway must be online. The report was uploaded; publishing again replaces it.";

                default:
                    return null;
            }
        }
    }
}
