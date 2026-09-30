using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace Flank.SsmsExtension
{
    internal static class SsmsUiHelper
    {
        public static void SetStatusBar(string text)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var statusBar =
                Package.GetGlobalService(typeof(SVsStatusbar))
                as IVsStatusbar;

            if (statusBar != null)
            {
                statusBar.SetText(text);
            }
        }

        public static void ShowMessage(
            string message,
            OLEMSGICON icon,
            string title = "Flank")
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            VsShellUtilities.ShowMessageBox(
                ServiceProvider.GlobalProvider,
                message,
                title,
                icon,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
