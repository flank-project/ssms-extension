using Flank.Excel;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Saving and opening a refreshable workbook, shared by the query editor
    /// and Object Explorer commands.
    /// </summary>
    internal static class ExcelWorkflow
    {
        public static async Task RunAsync(
            string sql,
            string server,
            string database,
            string defaultFileName,
            IReadOnlyList<string> parameterNames = null,
            IReadOnlyList<string> columnNames = null)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            string step = "Choosing output file";
            string outputPath = null;

            try
            {
                using (var dialog = new System.Windows.Forms.SaveFileDialog
                {
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    DefaultExt = "xlsx",
                    AddExtension = true,
                    FileName = defaultFileName,
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

                await Task.Run(() =>
                {
                    ExcelExporter.Generate(
                        outputPath,
                        sql,
                        server,
                        database,
                        parameterNames ?? Array.Empty<string>(),
                        columnNames);
                });

                await ThreadHelper.JoinableTaskFactory
                    .SwitchToMainThreadAsync();

                step = "Opening Excel workbook";

                SsmsUiHelper.SetStatusBar(
                    "Refreshable Excel created — opening Excel...");

                System.Diagnostics.Process.Start(outputPath);
            }
            catch (Exception ex)
            {
                await ThreadHelper.JoinableTaskFactory
                    .SwitchToMainThreadAsync();

                ShowFailure(
                    step,
                    ex,
                    server,
                    database,
                    outputPath);
            }
        }

        public static void ShowFailure(
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

            SsmsUiHelper.ShowMessage(
                message,
                OLEMSGICON.OLEMSGICON_CRITICAL,
                "Flank Error");
        }
    }
}
