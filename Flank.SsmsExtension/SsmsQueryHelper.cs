using Microsoft.VisualStudio.Shell;
using System;

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
                throw new Exception("Could not access the SSMS editor.");

            var doc = dte.ActiveDocument;

            if (doc == null)
                throw new Exception("Open a SQL query window first.");

            var selection = doc.Selection as EnvDTE.TextSelection;

            if (selection == null)
                throw new Exception("Could not read the current SQL editor.");

            if (!string.IsNullOrWhiteSpace(selection.Text))
                return selection.Text;

            var textDoc = doc.Object("TextDocument")
                as EnvDTE.TextDocument;

            if (textDoc == null)
                throw new Exception("Could not read the current SQL document.");

            var start = textDoc.StartPoint.CreateEditPoint();
            return start.GetText(textDoc.EndPoint);
        }
    }
}
