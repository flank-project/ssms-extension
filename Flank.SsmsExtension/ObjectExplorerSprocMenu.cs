using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Adds Flank's item to the Object Explorer right-click menu for stored procedures.
    /// </summary>
    /// <remarks>
    /// Object Explorer menus aren't VS command-table menus, so .vsct can't target them.
    /// Each node type's menu is a <see cref="HierarchyObject"/>, reachable from a node
    /// via its <see cref="IMenuHandler"/> service. We add our item the first time a
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
                objectExplorer.GetSelectedNodes(out int count, out INodeInformation[] nodes);

                if (count != 1 || nodes[0].UrnPath != StoredProcedureUrnPath)
                    return;

                if (nodes[0].GetService(typeof(IMenuHandler)) is HierarchyObject menu &&
                    extendedMenus.Add(menu))
                {
                    menu.AddChild(string.Empty, new SprocMenuItem());
                }
            }
            catch (Exception ex)
            {
                // Never let a menu hook break Object Explorer.
                System.Diagnostics.Debug.WriteLine($"Flank: Object Explorer hook failed: {ex}");
            }
        }

        private sealed class SprocMenuItem : ToolsMenuItemBase
        {
            public SprocMenuItem()
            {
                Text = "Make Self-Serve";
            }

            public override object Clone()
            {
                return new SprocMenuItem();
            }

            // SSMS swallows exceptions thrown from here, so report them ourselves.
            protected override void Invoke()
            {
                try
                {
                    InvokeCore();
                }
                catch (Exception ex)
                {
                    SsmsUiHelper.ShowMessage(
                        $"Make Self-Serve failed.\n\n{ex}",
                        OLEMSGICON.OLEMSGICON_CRITICAL);
                }
            }

            private void InvokeCore()
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                var sproc = StoredProcedureInfo.FromNode(Parent);

                if (sproc == null)
                {
                    SsmsUiHelper.ShowMessage(
                        "Could not read the selected stored procedure.",
                        OLEMSGICON.OLEMSGICON_WARNING);
                    return;
                }

                SsmsUiHelper.ShowMessage(
                    $"Server: {sproc.Server}\n" +
                    $"Database: {sproc.Database}\n" +
                    $"Procedure: [{sproc.Schema}].[{sproc.Name}]",
                    OLEMSGICON.OLEMSGICON_INFO);
            }
        }
    }

    internal sealed class StoredProcedureInfo
    {
        public string Server { get; private set; } = "";
        public string Database { get; private set; } = "";
        public string Schema { get; private set; } = "";
        public string Name { get; private set; } = "";

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
