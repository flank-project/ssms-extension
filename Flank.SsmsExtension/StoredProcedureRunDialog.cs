using Flank.Ssrs;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Runs a stored procedure to preview its results before creating an SSRS
    /// report from it (the report takes its columns from the run).
    /// </summary>
    internal sealed class StoredProcedureRunDialog : Form
    {
        private readonly string procedureName;
        private readonly string outputName;
        private readonly PreviewPanel preview;
        private readonly Button createButton;

        /// <param name="parameters">All of the procedure's parameters; OUTPUT ones get no input.</param>
        /// <param name="openConnection">Opens a connection to the procedure's database.</param>
        /// <param name="outputName">What's being created, for messages, e.g. "SSRS".</param>
        /// <param name="reviewed">From the parameter review page, if there was one.</param>
        /// <param name="runQuery">
        /// For a query rather than a procedure: runs it with the values, in
        /// parameter order. <paramref name="procedureName"/> is then just its name in messages.
        /// </param>
        /// <param name="canGoBack">
        /// There was a review page before this: a Back button closes the dialog
        /// with <see cref="DialogResult.Retry"/>, for the caller to show it again.
        /// </param>
        public StoredProcedureRunDialog(
            string procedureName,
            IReadOnlyList<SprocParameter> parameters,
            Func<IDbConnection> openConnection,
            string createText,
            string outputName,
            IReadOnlyList<ParameterSettings>? reviewed = null,
            Func<IReadOnlyList<object?>, IReadOnlyList<ResultSet>>? runQuery = null,
            bool canGoBack = false)
        {
            this.procedureName = procedureName;
            this.outputName = outputName;

            var inputs = parameters.Where(p => !p.IsOutput).ToList();
            var outputs = parameters.Where(p => p.IsOutput).ToList();

            Text = "Run " + procedureName;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            ClientSize = new Size(780, 600);
            MinimumSize = new Size(600, 450);

            Controls.Add(new Label
            {
                Text = runQuery == null ? "Run the procedure to preview its results." : "Run the query to preview its results.",
                Location = new Point(16, 12),
                Size = new Size(748, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            });

            preview = new PreviewPanel(inputs, new Size(748, ClientSize.Height - 40 - 52), reviewed)
            {
                Location = new Point(16, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Runner = runQuery ?? (values =>
                {
                    var named = inputs
                        .Select((p, i) => new KeyValuePair<string, object?>(p.Name, values[i]))
                        .ToList();

                    using (var connection = openConnection())
                    {
                        // Table-valued parameters go through the report's own call, which takes them as JSON.
                        return inputs.Any(p => p.IsTable)
                            ? StoredProcedureRunner.RunCall(connection, procedureName, parameters, reviewed, values, PreviewPanel.MaxRows)
                            : StoredProcedureRunner.Run(connection, procedureName, named, PreviewPanel.MaxRows, outputs);
                    }
                }),
            };
            Controls.Add(preview);

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(ClientSize.Width - 106, ClientSize.Height - 40),
                Width = 90,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };

            createButton = new Button
            {
                Text = createText,
                Enabled = false,
                Location = new Point(ClientSize.Width - 256, ClientSize.Height - 40),
                Width = 140,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };
            createButton.Click += (s, e) =>
            {
                if (!PreviewPanel.ConfirmFirstResultSetOnly(this, this.procedureName, preview.Results, this.outputName))
                    return;

                DialogResult = DialogResult.OK;
                Close();
            };

            preview.ResultsChanged += (s, e) =>
            {
                createButton.Enabled = preview.HasResults;
                AcceptButton = preview.HasResults ? createButton : preview.RunButton;
            };

            if (canGoBack)
            {
                var backButton = new Button
                {
                    Text = "< Back",
                    DialogResult = DialogResult.Retry,
                    Location = new Point(ClientSize.Width - 356, ClientSize.Height - 40),
                    Width = 90,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                };
                Controls.Add(backButton);
            }

            Controls.Add(cancelButton);
            Controls.Add(createButton);
            CancelButton = cancelButton;
            AcceptButton = preview.RunButton;
        }

        /// <summary>Typed values from the last run, in input-parameter order (null = default).</summary>
        public IReadOnlyList<object?> ParameterValues => preview.ParameterValues;

        /// <summary>Result sets from the last run.</summary>
        public IReadOnlyList<ResultSet> Results => preview.Results;

        /// <summary>When the last run happened.</summary>
        public DateTime RanAt => preview.RanAt;
    }
}
