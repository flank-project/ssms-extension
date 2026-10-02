using Microsoft.VisualStudio.Shell;
using System;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

using System.IO;

namespace Flank.SsmsExtension
{
    internal static class SsmsQueryHelper
    {
        public static string GetCurrentSql()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = Package.GetGlobalService(typeof(EnvDTE.DTE))
                as EnvDTE.DTE;

            if (dte == null)
                throw new InvalidOperationException(
                    "Could not access the SSMS editor.");

            var doc = dte.ActiveDocument;

            if (doc == null)
                throw new InvalidOperationException(
                    "Open a SQL query window first.");

            var selection = doc.Selection as EnvDTE.TextSelection;

            if (selection == null)
                throw new InvalidOperationException(
                    "Could not read the current SQL editor.");

            string sql;

            if (!string.IsNullOrWhiteSpace(selection.Text))
            {
                sql = selection.Text;
            }
            else
            {
                var textDoc = doc.Object("TextDocument")
                    as EnvDTE.TextDocument;

                if (textDoc == null)
                    throw new InvalidOperationException(
                        "Could not read the current SQL document.");

                var start = textDoc.StartPoint.CreateEditPoint();
                sql = start.GetText(textDoc.EndPoint);
            }

            if (string.IsNullOrWhiteSpace(sql))
                throw new InvalidOperationException(
                    "The current query is empty.");

            return sql;
        }
        public static IDbConnection GetCurrentConnection()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                "flank-connection-debug.txt");

            var log = new StringBuilder();

            try
            {
                log.AppendLine("=== Flank GetCurrentConnection ===");
                log.AppendLine("Time: " + DateTime.Now);
                log.AppendLine();

                var assemblies = AppDomain.CurrentDomain.GetAssemblies();

                log.AppendLine("Loaded assemblies containing 'Sql':");

                foreach (var assembly in assemblies
                    .Where(a =>
                        (a.GetName().Name ?? "")
                        .IndexOf("Sql", StringComparison.OrdinalIgnoreCase) >= 0)
                    .OrderBy(a => a.GetName().Name))
                {
                    log.AppendLine(
                        $"  {assembly.GetName().Name} | {assembly.FullName}");
                }

                log.AppendLine();
                log.AppendLine("STEP 1: Find ScriptEditorControl");

                var sqlEditorsAssembly = AppDomain.CurrentDomain
                    .GetAssemblies()
                    .FirstOrDefault(a =>
                        a.GetName().Name == "SqlEditors");

                var editorType = sqlEditorsAssembly?.GetType(
                    "Microsoft.SqlServer.Management.UI.VSIntegration.Editors.ScriptEditorControl",
                    throwOnError: false);

                if (editorType == null)
                {
                    log.AppendLine("FAILED: ScriptEditorControl not found.");
                    File.WriteAllText(logPath, log.ToString());

                    throw new InvalidOperationException(
                        "Could not access the SSMS SQL editor.");
                }

                log.AppendLine(
                    "SUCCESS: " + editorType.AssemblyQualifiedName);

                log.AppendLine();
                log.AppendLine("STEP 2: Find GetActiveScriptEditorControl");

                var method = editorType.GetMethod(
                    "GetActiveScriptEditorControl",
                    BindingFlags.Static |
                    BindingFlags.Public |
                    BindingFlags.NonPublic);

                if (method == null)
                {
                    log.AppendLine(
                        "FAILED: GetActiveScriptEditorControl not found.");

                    File.WriteAllText(logPath, log.ToString());

                    throw new InvalidOperationException(
                        "Could not access the active SSMS SQL editor.");
                }

                log.AppendLine("SUCCESS: " + method);

                log.AppendLine();
                log.AppendLine("STEP 3: Get active editor");

                var editor = method.Invoke(null, null);

                if (editor == null)
                {
                    log.AppendLine("FAILED: Active editor is null.");
                    File.WriteAllText(logPath, log.ToString());

                    throw new InvalidOperationException(
                        "Open a SQL query window first.");
                }

                log.AppendLine(
                    "SUCCESS: " + editor.GetType().AssemblyQualifiedName);

                log.AppendLine();
                log.AppendLine("STEP 4: Find m_connection");

                var connectionField = editor.GetType().GetField(
                    "m_connection",
                    BindingFlags.Instance |
                    BindingFlags.NonPublic);

                if (connectionField == null)
                {
                    log.AppendLine("FAILED: m_connection field not found.");

                    log.AppendLine();
                    log.AppendLine("Available fields:");

                    foreach (var field in editor.GetType().GetFields(
                        BindingFlags.Instance |
                        BindingFlags.Public |
                        BindingFlags.NonPublic))
                    {
                        log.AppendLine(
                            $"  {field.FieldType.FullName} {field.Name}");
                    }

                    File.WriteAllText(logPath, log.ToString());

                    throw new InvalidOperationException(
                        "Could not access the query connection.");
                }

                log.AppendLine(
                    "SUCCESS: " +
                    connectionField.FieldType.FullName +
                    " " +
                    connectionField.Name);

                log.AppendLine();
                log.AppendLine("STEP 5: Read m_connection");

                var rawConnection = connectionField.GetValue(editor);

                if (rawConnection == null)
                {
                    log.AppendLine("FAILED: m_connection is null.");
                    File.WriteAllText(logPath, log.ToString());

                    throw new InvalidOperationException(
                        "The current query window is not connected.");
                }

                log.AppendLine(
                    "Runtime type: " +
                    rawConnection.GetType().AssemblyQualifiedName);

                var connection = rawConnection as IDbConnection;

                if (connection == null)
                {
                    log.AppendLine(
                        "FAILED: Connection does not implement IDbConnection.");

                    File.WriteAllText(logPath, log.ToString());

                    throw new InvalidOperationException(
                        "Could not access the query connection.");
                }

                log.AppendLine("IDbConnection cast: SUCCESS");
                log.AppendLine("Connection state: " + connection.State);

                if (connection.State != ConnectionState.Open)
                {
                    log.AppendLine("FAILED: Connection is not open.");
                    File.WriteAllText(logPath, log.ToString());

                    throw new InvalidOperationException(
                        "The current query window is not connected.");
                }

                log.AppendLine();
                log.AppendLine("SUCCESS");

                File.WriteAllText(logPath, log.ToString());

                return connection;
            }
            catch
            {
                // Preserve whatever diagnostics we have even for an
                // unexpected reflection exception.
                try
                {
                    File.WriteAllText(logPath, log.ToString());
                }
                catch
                {
                }

                throw;
            }
        }
    }
}
