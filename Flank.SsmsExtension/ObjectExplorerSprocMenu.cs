using Flank.Ssrs;
using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Adds Make Self-Serve to the Object Explorer right-click menu for stored procedures.
    /// </summary>
    /// <remarks>
    /// Object Explorer menus aren't VS command-table menus, so .vsct can't target them.
    /// Each node type's menu is a <see cref="HierarchyObject"/>, reachable from a node
    /// via its <see cref="IMenuHandler"/> service. We add our submenu the first time a
    /// stored procedure is selected (right-clicking a node selects it first).
    /// </remarks>
    internal sealed class ObjectExplorerSprocMenu
    {
        private const string StoredProcedureUrnPath = "Server/Database/StoredProcedure";

        private readonly IObjectExplorerService objectExplorer;

        // A menu is shared by every node of its type; remember which ones already have our item.
        private readonly HashSet<HierarchyObject> extendedMenus = new HashSet<HierarchyObject>();

        private ObjectExplorerSprocMenu(IObjectExplorerService objectExplorer)
        {
            this.objectExplorer = objectExplorer;
        }

        public static ObjectExplorerSprocMenu Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(
                package.DisposalToken);

            var objectExplorer =
                await package.GetServiceAsync(typeof(IObjectExplorerService))
                as IObjectExplorerService;

            var contextService =
                await package.GetServiceAsync(typeof(IContextService))
                as IContextService;

            if (objectExplorer == null || contextService == null)
                return;

            Instance = new ObjectExplorerSprocMenu(objectExplorer);
            contextService.ActionContext.CurrentContextChanged += Instance.OnContextChanged;
        }

        private void OnContextChanged(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var node = GetSelectedStoredProcedure();

                if (node?.GetService(typeof(IMenuHandler)) is HierarchyObject menu &&
                    extendedMenus.Add(menu))
                {
                    menu.AddChild(string.Empty, new SelfServeSubmenu(this));
                }
            }
            catch (Exception ex)
            {
                // Never let a menu hook break Object Explorer.
                System.Diagnostics.Debug.WriteLine($"Flank: Object Explorer hook failed: {ex}");
            }
        }

        private INodeInformation GetSelectedStoredProcedure()
        {
            objectExplorer.GetSelectedNodes(out int count, out INodeInformation[] nodes);

            return count == 1 && nodes[0].UrnPath == StoredProcedureUrnPath
                ? nodes[0]
                : null;
        }

        private void RunForSelection(Func<INodeInformation, StoredProcedureInfo, Task> action)
        {
            ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                var node = GetSelectedStoredProcedure();
                var sproc = StoredProcedureInfo.FromNode(node);

                if (sproc == null)
                {
                    SsmsUiHelper.ShowMessage(
                        "Select a stored procedure in Object Explorer first.",
                        OLEMSGICON.OLEMSGICON_INFO);
                    return;
                }

                await action(node, sproc);
            });
        }

        private static async Task CreateExcelAsync(INodeInformation node, StoredProcedureInfo sproc)
        {
            string step = "Reading stored procedure";

            try
            {
                using (var connection = OpenConnection(node, sproc.Database))
                {
                    int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(
                        sproc.QualifiedName, connection);

                    var parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(
                        objectId, connection);

                    if (parameters.Count > 0)
                    {
                        SsmsUiHelper.ShowMessage(
                            $"{sproc.QualifiedName} has parameters " +
                            $"({string.Join(", ", parameters.Select(p => p.Name))}).\n\n" +
                            "Excel workbooks only support stored procedures without " +
                            "parameters for now. SSRS Report supports parameters.",
                            OLEMSGICON.OLEMSGICON_INFO);
                        return;
                    }

                    // Fails with a clear message if there's no result set to load.
                    SqlMetadataDiscovery.DiscoverStoredProcedureColumns(
                        objectId, sproc.QualifiedName, connection);
                }
            }
            catch (Exception ex)
            {
                ExcelWorkflow.ShowFailure(step, ex, sproc.Server, sproc.Database, null);
                return;
            }

            await ExcelWorkflow.RunAsync(
                $"EXEC {sproc.QualifiedName}",
                sproc.Server,
                sproc.Database,
                sproc.Name + ".xlsx");
        }

        private static Task CreateSsrsReportAsync(INodeInformation node, StoredProcedureInfo sproc)
        {
            return SsrsWorkflow.RunAsync(
                $"EXEC {sproc.QualifiedName}",
                sqlReadOnly: true,
                defaultReportName: sproc.Name,
                analyzingStatus: "Analyzing stored procedure...",
                prepare: reportOptions =>
                {
                    using (var connection = OpenConnection(node, sproc.Database))
                    {
                        int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(
                            sproc.QualifiedName, connection);

                        var parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(
                            objectId, connection);

                        var columns = SqlMetadataDiscovery.DiscoverStoredProcedureColumns(
                            objectId, sproc.QualifiedName, connection);

                        return client => client.CreateReportFromStoredProcedure(
                            sproc.QualifiedName,
                            parameters,
                            columns,
                            reportOptions.DataSource,
                            reportOptions.Folder,
                            reportOptions.ReportName);
                    }
                });
        }

        /// <summary>Opens a connection to the node's server, in the procedure's database.</summary>
        private static IDbConnection OpenConnection(INodeInformation node, string database)
        {
            if (!(node.Connection is SqlConnectionInfo info))
                throw new InvalidOperationException("Could not read the Object Explorer connection.");

            var copy = info.Copy();
            copy.DatabaseName = database;

            var connection = copy.CreateConnectionObject();
            connection.Open();
            return connection;
        }

        /// <summary>
        /// The Make Self-Serve submenu. Object Explorer asks menu children that are
        /// themselves menu handlers for their WinForms items and adds them as-is.
        /// Must not implement IMenuItem, or it's added as a plain item instead.
        /// </summary>
        private sealed class SelfServeSubmenu : IMenuHandler, IWinformsMenuHandler, ICloneable
        {
            private readonly ObjectExplorerSprocMenu owner;

            public SelfServeSubmenu(ObjectExplorerSprocMenu owner)
            {
                this.owner = owner;
            }

            public INodeInformation Parent { get; set; }

            public CommandID ContextMenuID => null;

            public event EventHandler OnRefresh
            {
                add { }
                remove { }
            }

            public ToolStripItem[] GetMenuItems()
            {
                var submenu = new ToolStripMenuItem("Make Self-Serve");

                submenu.DropDownItems.Add(
                    "Excel (Power Query)",
                    null,
                    (s, e) => owner.RunForSelection(CreateExcelAsync));

                submenu.DropDownItems.Add(
                    "SSRS Report",
                    null,
                    (s, e) => owner.RunForSelection(CreateSsrsReportAsync));

                return new ToolStripItem[] { submenu };
            }

            public object Clone() => new SelfServeSubmenu(owner);

            public void FillMenuCommands(IMenuCommandService menuCommands) { }

            public void UpdateMenuCommandsStatus(MenuCommand menuCommand) { }

            public void DoDefaultAction() { }

            public void InvokeProperties() { }
        }
    }

    internal sealed class StoredProcedureInfo
    {
        public string Server { get; private set; } = "";
        public string Database { get; private set; } = "";
        public string Schema { get; private set; } = "";
        public string Name { get; private set; } = "";

        /// <summary>[schema].[name], quoted for T-SQL.</summary>
        public string QualifiedName => $"{QuoteName(Schema)}.{QuoteName(Name)}";

        /// <summary>
        /// Reads a stored procedure node's URN, e.g.
        /// Server[@Name='srv']/Database[@Name='db']/StoredProcedure[@Name='p' and @Schema='dbo'].
        /// </summary>
        public static StoredProcedureInfo? FromNode(INodeInformation? node)
        {
            if (node == null)
                return null;

            string urn = node.Context;

            // The URN's server is the server's own name (on Azure SQL, without
            // ".database.windows.net"); the connection has the address to connect to.
            string? server = node.Connection?.ServerName ?? UrnAttribute(urn, "Server", "Name");
            string? database = UrnAttribute(urn, "Database", "Name");
            string? name = UrnAttribute(urn, "StoredProcedure", "Name");
            string? schema = UrnAttribute(urn, "StoredProcedure", "Schema");

            if (server == null || database == null || name == null || schema == null)
                return null;

            return new StoredProcedureInfo
            {
                Server = server,
                Database = database,
                Schema = schema,
                Name = name,
            };
        }

        private static string QuoteName(string name) => "[" + name.Replace("]", "]]") + "]";

        // URN values are single-quoted, with embedded quotes doubled.
        private static string? UrnAttribute(string urn, string element, string attribute)
        {
            var match = Regex.Match(
                urn,
                $@"(?:^|/){element}\[[^\]]*?@{attribute}='((?:[^']|'')*)'");

            return match.Success ? match.Groups[1].Value.Replace("''", "'") : null;
        }
    }
}
