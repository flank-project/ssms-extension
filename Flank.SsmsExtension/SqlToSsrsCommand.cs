using Flank.Ssrs;
using Flank.Ssrs;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;
using System.Linq;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SqlToSsrsCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0110;

        public static readonly Guid CommandSet =
            new Guid("4aefd085-bfeb-4d3d-8199-201b0c86abc0");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlToSsrsCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SqlToSsrsCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static SqlToSsrsCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in SqlToSsrsCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SqlToSsrsCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.JoinableTaskFactory.Run(async () =>
            {
                await ExecuteAsync();
            });
        }

        private async Task ExecuteAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            string sql;

            try
            {
                sql = SsmsQueryHelper.GetCurrentSql();
            }
            catch (InvalidOperationException ex)
            {
                SsmsUiHelper.ShowMessage(
                    ex.Message,
                    OLEMSGICON.OLEMSGICON_INFO);

                return;
            }

            try
            {
                // Hardcoded for now — just proving the SSMS -> SSRS path.
                var client = new SsrsClient(
                    "http://localhost/ReportServer",
                    "http://localhost/Reports");

                //DumpLikelyQueryTypes();

                //foreach (System.Windows.Forms.Form form
                //    in System.Windows.Forms.Application.OpenForms)
                //{
                //    FindSqlEditorControl(form);
                //}

                //DumpTypeMembers(
                //    "Microsoft.SqlServer.Management.UI.VSIntegration.Editors.ScriptEditorControl");

                //DumpTypeMembers(
                //    "Microsoft.SqlServer.Management.UI.VSIntegration.Editors.ScriptFactory");

                //            DumpTypeMethods(
                //"Microsoft.SqlServer.Management.UI.VSIntegration.Editors.SqlScriptEditorControl");

                //            DumpTypeMethods(
                //                "Microsoft.SqlServer.Management.UI.VSIntegration.Editors.ScriptEditorControl");

                //            DumpTypeMethods(
                //                "Microsoft.SqlServer.Management.UI.VSIntegration.Editors.ScriptFactory");

                var editorType = AppDomain.CurrentDomain
                    .GetAssemblies()
                    .SelectMany(a =>
                    {
                        try { return a.GetTypes(); }
                        catch { return Type.EmptyTypes; }
                    })
                    .FirstOrDefault(t =>
                        t.FullName ==
                        "Microsoft.SqlServer.Management.UI.VSIntegration.Editors.ScriptEditorControl");

                Debug.WriteLine($"Type: {editorType?.FullName ?? "null"}");

                var method = editorType.GetMethod(
                    "GetActiveScriptEditorControl",
                    BindingFlags.Static |
                    BindingFlags.Public |
                    BindingFlags.NonPublic);

                var editor = method?.Invoke(null, null);

                Debug.WriteLine(
                    $"Editor: {editor?.GetType().FullName ?? "null"}");

                var connectionField = editor.GetType().GetField(
                    "m_connection",
                    BindingFlags.Instance |
                    BindingFlags.NonPublic);

                var connectionTest =
                    connectionField?.GetValue(editor) as System.Data.IDbConnection;

                Debug.WriteLine(
                    $"Connection type: {connectionTest?.GetType().FullName ?? "null"}");

                Debug.WriteLine(
                    $"Connection state: {connectionTest?.State}");
                //DumpTypeMembers("Microsoft.SqlServer.Management.UI.VSIntegration.Editors.SqlScriptEditorControl");

                //var scriptFactory = ServiceCache.ScriptFactory;

                //DumpInterestingMembers(
                //    scriptFactory,
                //    "ScriptFactory");
                var connection = SsmsQueryHelper.GetCurrentConnection();


                string reportUrl = await Task.Run(() =>
                {
                    return client.CreateReport(
                        sql,
                        connection,
                        "/HardcodedTest",
                        "/TestFolder",
                        "SSMS Query Test");
                });

                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                System.Diagnostics.Process.Start(reportUrl);
            }
            catch (Exception ex)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                SsmsUiHelper.ShowMessage(
                    "Flank couldn't create the SSRS report.\n\n" +
                    ex.Message,
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    "Flank Error");
            }
        }

        private static void DumpInterestingMembers(
    object obj,
    string path = "root",
    int depth = 0,
    int maxDepth = 4)
        {
            if (obj == null || depth > maxDepth)
                return;

            var type = obj.GetType();

            Debug.WriteLine(
                $"{new string(' ', depth * 2)}{path} [{type.FullName}]");

            var flags =
                BindingFlags.Instance |
                BindingFlags.Public |
                BindingFlags.NonPublic;

            foreach (var field in type.GetFields(flags))
            {
                try
                {
                    var value = field.GetValue(obj);

                    if (IsInteresting(field.Name, field.FieldType))
                    {
                        Debug.WriteLine(
                            $"{new string(' ', (depth + 1) * 2)}" +
                            $"{field.Name} = {value?.GetType().FullName ?? "null"}");
                    }

                    if (ShouldExplore(field.Name, field.FieldType, value))
                    {
                        DumpInterestingMembers(
                            value,
                            path + "." + field.Name,
                            depth + 1,
                            maxDepth);
                    }
                }
                catch
                {
                    // Ignore members that can't safely be inspected.
                }
            }

            foreach (var property in type.GetProperties(flags))
            {
                if (property.GetIndexParameters().Length > 0)
                    continue;

                try
                {
                    var value = property.GetValue(obj, null);

                    if (IsInteresting(property.Name, property.PropertyType))
                    {
                        Debug.WriteLine(
                            $"{new string(' ', (depth + 1) * 2)}" +
                            $"{property.Name} = {value?.GetType().FullName ?? "null"}");
                    }

                    if (ShouldExplore(property.Name, property.PropertyType, value))
                    {
                        DumpInterestingMembers(
                            value,
                            path + "." + property.Name,
                            depth + 1,
                            maxDepth);
                    }
                }
                catch
                {
                    // Some SSMS properties throw when inspected.
                }
            }
        }

        private static bool IsInteresting(string name, Type type)
        {
            var text = (name + " " + type.FullName).ToLowerInvariant();

            return
                text.Contains("connection") ||
                text.Contains("query") ||
                text.Contains("execution") ||
                text.Contains("sqlconnection") ||
                text.Contains("serverconnection");
        }

        private static bool ShouldExplore(
            string name,
            Type type,
            object value)
        {
            if (value == null)
                return false;

            if (!IsInteresting(name, type))
                return false;

            if (type.IsPrimitive ||
                type == typeof(string) ||
                type.IsEnum)
                return false;

            return true;
        }

        private static void DumpLikelyQueryTypes()
        {
            var terms = new[]
            {
        "QueryManager",
        "ExecutionManager",
        "SqlScriptEditor",
        "ScriptEditorControl",
        "QueryExecution"
    };

            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                Type[] types;

                try
                {
                    types = assembly.GetTypes();
                }
                catch (ReflectionTypeLoadException ex)
                {
                    types = ex.Types.Where(t => t != null).ToArray();
                }
                catch
                {
                    continue;
                }

                foreach (var type in types)
                {
                    if (terms.Any(term =>
                        type.FullName.IndexOf(
                            term,
                            StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        Debug.WriteLine(
                            $"{assembly.GetName().Name}: {type.FullName}");
                    }
                }
            }
        }

        private static void DumpTypeMembers(string fullTypeName)
        {
            var type = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(a =>
                {
                    try { return a.GetTypes(); }
                    catch (ReflectionTypeLoadException ex)
                    {
                        return ex.Types.Where(t => t != null);
                    }
                    catch
                    {
                        return Enumerable.Empty<Type>();
                    }
                })
                .FirstOrDefault(t => t.FullName == fullTypeName);

            if (type == null)
            {
                Debug.WriteLine("TYPE NOT FOUND");
                return;
            }

            Debug.WriteLine($"TYPE: {type.FullName}");

            var flags =
                BindingFlags.Instance |
                BindingFlags.Static |
                BindingFlags.Public |
                BindingFlags.NonPublic;

            foreach (var field in type.GetFields(flags))
                Debug.WriteLine(
                    $"FIELD: {field.Name} : {field.FieldType.FullName}");

            foreach (var property in type.GetProperties(flags))
                Debug.WriteLine(
                    $"PROPERTY: {property.Name} : {property.PropertyType.FullName}");
        }
        private static void FindSqlEditorControl(
    System.Windows.Forms.Control control,
    int depth = 0)
        {
            if (control == null)
                return;

            var type = control.GetType();

            if (type.FullName ==
                "Microsoft.SqlServer.Management.UI.VSIntegration.Editors.SqlScriptEditorControl")
            {
                Debug.WriteLine("FOUND SQL EDITOR CONTROL!");
                Debug.WriteLine($"Type: {type.FullName}");

                var connectionField = type.GetField(
                    "m_connection",
                    BindingFlags.Instance |
                    BindingFlags.NonPublic);

                var connection = connectionField?.GetValue(control);

                Debug.WriteLine(
                    $"m_connection runtime type: " +
                    $"{connection?.GetType().FullName ?? "null"}");

                return;
            }

            foreach (System.Windows.Forms.Control child in control.Controls)
            {
                FindSqlEditorControl(child, depth + 1);
            }
        }

        private static void DumpTypeMethods(string fullTypeName)
        {
            var type = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(a =>
                {
                    try { return a.GetTypes(); }
                    catch (ReflectionTypeLoadException ex)
                    {
                        return ex.Types.Where(t => t != null);
                    }
                    catch
                    {
                        return Enumerable.Empty<Type>();
                    }
                })
                .FirstOrDefault(t => t.FullName == fullTypeName);

            if (type == null)
            {
                Debug.WriteLine("TYPE NOT FOUND");
                return;
            }

            Debug.WriteLine($"METHODS: {type.FullName}");

            var flags =
                BindingFlags.Instance |
                BindingFlags.Static |
                BindingFlags.Public |
                BindingFlags.NonPublic;

            foreach (var method in type.GetMethods(flags)
                .OrderBy(m => m.Name))
            {
                var parameters = string.Join(
                    ", ",
                    method.GetParameters()
                        .Select(p => $"{p.ParameterType.Name} {p.Name}"));

                Debug.WriteLine(
                    $"{method.ReturnType.FullName} " +
                    $"{method.Name}({parameters})");
            }
        }
    }
}
