using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Flank.Excel;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SqlToExcelPowerQueryCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("4aefd085-bfeb-4d3d-8199-201b0c86abc0");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlToExcelPowerQueryCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SqlToExcelPowerQueryCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static SqlToExcelPowerQueryCommand Instance
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
            // Switch to the main thread - the call to AddCommand in Command1's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SqlToExcelPowerQueryCommand(package, commandService);
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
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;

            VsShellUtilities.ShowMessageBox(
                this.package,
                dte == null ? "DTE is NULL" : "DTE exists!",
                "Flank Debug",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            var doc = dte.ActiveDocument;

            if (doc == null)
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "Open a query window first.",
                    "Flank",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                return;
            }

            var selection = (EnvDTE.TextSelection)doc.Selection;

            string sql;

            if (!string.IsNullOrWhiteSpace(selection.Text))
            {
                sql = selection.Text;
            }
            else
            {
                var textDoc = (EnvDTE.TextDocument)
                    doc.Object("TextDocument");

                var start = textDoc.StartPoint.CreateEditPoint();
                sql = start.GetText(textDoc.EndPoint);
            }


            VsShellUtilities.ShowMessageBox(
                this.package,
                sql,
                "Flank SQL",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            var scriptFactory = ServiceCache.ScriptFactory;
            var connection = scriptFactory.CurrentlyActiveWndConnectionInfo;
            var info = connection.UIConnectionInfo;
            if (info == null)
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "No active connection.",
                    "Flank",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                return;
            }
            var details =
                $"ServerName: {info.ServerName}\n" +
                $"DisplayName: {info.DisplayName}\n" +
                $"OtherParams: {info.OtherParams}\n\n" +
                "AdvancedOptions:\n";

            foreach (string key in info.AdvancedOptions.AllKeys)
            {
                details += $"{key} = {info.AdvancedOptions[key]}\n";
            }

            VsShellUtilities.ShowMessageBox(
                this.package,
                details,
                "Flank Connection",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            string server = info.ServerName;
            string database = info.AdvancedOptions["DATABASE"];

            using var dialog = new System.Windows.Forms.SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                AddExtension = true,
                FileName = "report.xlsx",
                Title = "Save as refreshable Excel"
            };

            if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            string outputPath = dialog.FileName;

            ExcelExporter.Generate(
                outputPath,
                server,
                database,
                sql);
            System.Diagnostics.Process.Start(outputPath);
        }
    }
}
