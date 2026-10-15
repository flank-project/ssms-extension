using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using Task = System.Threading.Tasks.Task;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Make Self-Serve → Feedback... in the query editor.
    /// </summary>
    internal static class FeedbackCommand
    {
        public const int CommandId = 0x0120;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory
                .SwitchToMainThreadAsync(package.DisposalToken);

            var commandService =
                await package.GetServiceAsync(typeof(IMenuCommandService))
                    as OleMenuCommandService;

            commandService.AddCommand(new MenuCommand(
                (sender, e) => Feedback.ShowDialog(),
                new CommandID(SqlToSsrsCommand.CommandSet, CommandId)));
        }
    }
}
