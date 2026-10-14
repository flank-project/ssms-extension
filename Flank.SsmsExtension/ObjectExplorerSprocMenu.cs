using Flank.Excel;
using Flank.Ssrs;
using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
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
                    var submenu = new SelfServeSubmenu(this);
                    menu.AddChild(string.Empty, submenu);
                    MoveToTop(menu, submenu);
                }
            }
            catch (Exception ex)
            {
                // Never let a menu hook break Object Explorer.
                System.Diagnostics.Debug.WriteLine($"Flank: Object Explorer hook failed: {ex}");
            }
        }

        /// <summary>
        /// AddChild appends, and the menu is built in child order. Move our submenu
        /// to the front of the menu's private child list so it's the first entry.
        /// If SSMS changes that field, the submenu stays where AddChild put it.
        /// </summary>
        private static void MoveToTop(HierarchyObject menu, object child)
        {
            for (var type = menu.GetType(); type != null; type = type.BaseType)
            {
                var field = type.GetField(
                    "menuItemsInOrder",
                    BindingFlags.Instance | BindingFlags.NonPublic);

                if (field?.GetValue(menu) is ArrayList items && items.Contains(child))
                {
                    items.Remove(child);
                    items.Insert(0, child);
                    return;
                }
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
            IReadOnlyList<SprocParameter> parameters;

            try
            {
                using (var connection = OpenConnection(node, sproc.Database))
                {
                    int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(
                        sproc.QualifiedName, connection);

                    parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(
                        objectId, connection);

                    var outputs = parameters.Where(p => p.IsOutput).ToList();

                    if (outputs.Count > 0)
                    {
                        SsmsUiHelper.ShowMessage(
                            $"{sproc.QualifiedName} has OUTPUT parameters " +
                            $"({string.Join(", ", outputs.Select(p => p.Name))}).\n\n" +
                            "Excel workbooks don't support OUTPUT parameters yet.",
                            OLEMSGICON.OLEMSGICON_INFO);
                        return;
                    }

                }
            }
            catch (Exception ex)
            {
                ExcelWorkflow.ShowFailure(step, ex, sproc.Server, sproc.Database, null);
                return;
            }

            // Run it first: the workbook starts with these values and rows.
            IReadOnlyList<object?> values;
            ResultSet result;
            DateTime ranAt;

            using (var dialog = new StoredProcedureRunDialog(
                sproc.QualifiedName,
                parameters,
                () => OpenConnection(node, sproc.Database),
                "Create workbook"))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                values = dialog.ParameterValues;
                result = dialog.Results[0];
                ranAt = dialog.RanAt;
            }

            // Parameters become cells the user fills in (names in column A,
            // values in column B); the query passes them as @name.
            var names = parameters.Select(p => p.Name.TrimStart('@')).ToList();

            string sql = names.Count == 0
                ? $"EXEC {sproc.QualifiedName}"
                : $"EXEC {sproc.QualifiedName} " +
                  string.Join(", ", names.Select(n => $"@{n} = @{n}"));

            var request = new WorkbookRequest(sql, sproc.Server, sproc.Database)
            {
                ParameterNames = names,
                ParameterValues = values,
                ColumnNames = result.Columns,
                Rows = result.Rows,
                RowsFetchedAt = ranAt,
            };

            await ExcelWorkflow.RunAsync(request, sproc.Name + ".xlsx");
        }

        private static async Task CreateSsrsReportAsync(INodeInformation node, StoredProcedureInfo sproc)
        {
            IReadOnlyList<SprocParameter> parameters;

            try
            {
                using (var connection = OpenConnection(node, sproc.Database))
                {
                    int objectId = SqlMetadataDiscovery.GetStoredProcedureObjectId(
                        sproc.QualifiedName, connection);

                    parameters = SqlMetadataDiscovery.DiscoverStoredProcedureParameters(
                        objectId, connection);
                }
            }
            catch (Exception ex)
            {
                SsmsUiHelper.ShowMessage(
                    "Flank couldn't read the stored procedure.\n\n" + ex.Message,
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    "Flank Error");
                return;
            }

            // Run it first: the values become the report parameters' defaults,
            // so the report shows data when opened.
            var inputs = parameters.Where(p => !p.IsOutput).ToList();
            var defaults = new Dictionary<string, object?>();
            ResultSet result;

            using (var dialog = new StoredProcedureRunDialog(
                sproc.QualifiedName,
                inputs,
                () => OpenConnection(node, sproc.Database),
                "Continue"))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                for (int i = 0; i < inputs.Count; i++)
                    defaults[inputs[i].Name] = dialog.ParameterValues[i];

                result = dialog.Results[0];
            }

            var columns = result.Columns
                .Select(name => new ReportColumn { Name = name })
                .ToList();

            await SsrsWorkflow.RunAsync(
                sql: null,
                storedProcedure: sproc.QualifiedName,
                defaultReportName: sproc.Name,
                analyzingStatus: "Preparing report...",
                prepare: reportOptions => client => client.CreateReportFromStoredProcedure(
                    sproc.QualifiedName,
                    parameters,
                    columns,
                    reportOptions.DataSource,
                    reportOptions.Folder,
                    reportOptions.ReportName,
                    defaults));
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
                    MenuIcons.Excel,
                    (s, e) => owner.RunForSelection(CreateExcelAsync));

                submenu.DropDownItems.Add(
                    "SSRS Report",
                    MenuIcons.Ssrs,
                    (s, e) => owner.RunForSelection(CreateSsrsReportAsync));

                // We're first in the menu (see MoveToTop); separate us from SSMS's items.
                return new ToolStripItem[] { submenu, new ToolStripSeparator() };
            }

            public object Clone() => new SelfServeSubmenu(owner);

            public void FillMenuCommands(IMenuCommandService menuCommands) { }

            public void UpdateMenuCommandsStatus(MenuCommand menuCommand) { }

            public void DoDefaultAction() { }

            public void InvokeProperties() { }
        }
    }

    /// <summary>
    /// The query editor menus' icons, from the same 16px strip the .vsct uses
    /// (Excel, then SSRS), for the WinForms Object Explorer submenu.
    /// </summary>
    internal static class MenuIcons
    {
        private static readonly Lazy<Image?[]> icons = new Lazy<Image?[]>(Load);

        public static Image? Excel => icons.Value[0];

        public static Image? Ssrs => icons.Value[1];

        private static Image?[] Load()
        {
            try
            {
                using (var stream = typeof(MenuIcons).Assembly.GetManifestResourceStream("Flank.MenuIcons.png"))
                using (var strip = new Bitmap(stream))
                {
                    return new Image?[] { Slice(strip, 0), Slice(strip, 1) };
                }
            }
            catch (Exception ex)
            {
                // A missing icon shouldn't stop the menu working.
                System.Diagnostics.Debug.WriteLine($"Flank: could not load menu icons: {ex}");
                return new Image?[2];
            }
        }

        // Copies the pixels, so the icon doesn't depend on the strip's stream.
        private static Image Slice(Bitmap strip, int index)
        {
            var icon = new Bitmap(16, 16);

            using (var graphics = Graphics.FromImage(icon))
            {
                graphics.DrawImage(
                    strip,
                    new Rectangle(0, 0, 16, 16),
                    new Rectangle(index * 16, 0, 16, 16),
                    GraphicsUnit.Pixel);
            }

            return icon;
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
