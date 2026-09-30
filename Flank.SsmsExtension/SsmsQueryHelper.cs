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
                throw new InvalidOperationException(
                    "Could not access the SSMS editor.");

            var doc = dte.ActiveDocument;

            if (doc == null)
                throw new InvalidOperationException(
                    "Open a SQL query window first.");

            var selection = doc.Selection as EnvDTE.TextSelection;

            if (selection == null)
                throw new InvalidOperationException(
                    "Could not read the current SQL editor.");

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
                    throw new InvalidOperationException(
                        "Could not read the current SQL document.");

                var start = textDoc.StartPoint.CreateEditPoint();
                sql = start.GetText(textDoc.EndPoint);
            }

            if (string.IsNullOrWhiteSpace(sql))
                throw new InvalidOperationException(
                    "The current query is empty.");

            return sql;
        }
    }
}
