using Flank.Excel;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    internal sealed class SqlToExcelPowerQueryCommand
    {
        public const int CommandId = 0x0100;

        public static readonly Guid CommandSet =
            new Guid("4aefd085-bfeb-4d3d-8199-201b0c86abc0");

        private readonly AsyncPackage package;

        private SqlToExcelPowerQueryCommand(
            AsyncPackage package,
            OleMenuCommandService commandService)
        {
            this.package = package
                ?? throw new ArgumentNullException(nameof(package));

            commandService = commandService
                ?? throw new ArgumentNullException(nameof(commandService));

            var commandId = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, commandId);

            commandService.AddCommand(menuItem);
        }

        public static SqlToExcelPowerQueryCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(
                package.DisposalToken);

            var commandService =
                await package.GetServiceAsync(typeof(IMenuCommandService))
                as OleMenuCommandService;

            Instance = new SqlToExcelPowerQueryCommand(
                package,
                commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string step = "Reading query";
            string server = null;
            string database = null;
            string outputPath = null;

            try
            {
                var dte = Package.GetGlobalService(typeof(EnvDTE.DTE))
                    as EnvDTE.DTE;

                if (dte == null)
                    throw new Exception("Could not access the SSMS editor.");

                var doc = dte.ActiveDocument;

                if (doc == null)
                {
                    ShowMessage(
                        "Open a SQL query window first.",
                        OLEMSGICON.OLEMSGICON_INFO);

                    return;
                }

                var selection = doc.Selection as EnvDTE.TextSelection;

                if (selection == null)
                    throw new Exception("Could not read the current SQL editor.");

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
                        throw new Exception(
                            "Could not read the current SQL document.");

                    var start = textDoc.StartPoint.CreateEditPoint();
                    sql = start.GetText(textDoc.EndPoint);
                }

                if (string.IsNullOrWhiteSpace(sql))
                {
                    ShowMessage(
                        "The current query is empty.",
                        OLEMSGICON.OLEMSGICON_INFO);

                    return;
                }

                step = "Reading SQL Server connection";

                var scriptFactory = ServiceCache.ScriptFactory;

                if (scriptFactory == null)
                    throw new Exception(
                        "Could not access the SSMS connection service.");

                var connection =
                    scriptFactory.CurrentlyActiveWndConnectionInfo;

                var info = connection.UIConnectionInfo;

                if (info == null)
                    throw new Exception(
                        "No active SQL Server connection was found.");

                server = info.ServerName;
                database = info.AdvancedOptions["DATABASE"];

                if (string.IsNullOrWhiteSpace(server))
                    throw new Exception(
                        "Could not determine the SQL Server name.");

                if (string.IsNullOrWhiteSpace(database))
                    throw new Exception(
                        "Could not determine the database name.");

                step = "Choosing output file";

                using (var dialog = new System.Windows.Forms.SaveFileDialog
                {
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    DefaultExt = "xlsx",
                    AddExtension = true,
                    FileName = "report.xlsx",
                    Title = "Save as refreshable Excel"
                })
                {
                    if (dialog.ShowDialog() !=
                        System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }

                    outputPath = dialog.FileName;
                }

                step = "Creating Excel workbook";

                var waitDialogFactory =
                    Package.GetGlobalService(
                        typeof(SVsThreadedWaitDialogFactory))
                    as IVsThreadedWaitDialogFactory;

                IVsThreadedWaitDialog2 waitDialog = null;

                try
                {
                    if (waitDialogFactory != null)
                    {
                        waitDialogFactory.CreateInstance(
                            out waitDialog);

                        waitDialog.StartWaitDialog(
                            "Flank",
                            "Creating refreshable Excel workbook...",
                            null,
                            null,
                            null,
                            0,
                            false,
                            true);
                    }

                    ExcelExporter.Generate(
                        outputPath,
                        sql,
                        server,
                        database);
                }
                finally
                {
                    if (waitDialog != null)
                    {
                        int canceled;
                        waitDialog.EndWaitDialog(out canceled);
                    }
                }

                step = "Opening Excel workbook";

                System.Diagnostics.Process.Start(outputPath);
            }
            catch (Exception ex)
            {
                ShowFailure(
                    step,
                    ex,
                    server,
                    database,
                    outputPath);
            }
        }

        private void ShowFailure(
            string step,
            Exception exception,
            string server,
            string database,
            string outputPath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var message =
                "Flank couldn't create the refreshable Excel workbook.\n\n" +
                $"Step: {step}\n" +
                $"Error: {exception.Message}\n\n" +
                $"Server: {server ?? "(unknown)"}\n" +
                $"Database: {database ?? "(unknown)"}\n" +
                $"Output: {outputPath ?? "(not chosen yet)"}\n\n" +
                "Please screenshot this message when reporting the issue.";

            ShowMessage(
                message,
                OLEMSGICON.OLEMSGICON_CRITICAL,
                "Flank Error");
        }

        private void ShowMessage(
            string message,
            OLEMSGICON icon,
            string title = "Flank")
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            VsShellUtilities.ShowMessageBox(
                package,
                message,
                title,
                icon,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}