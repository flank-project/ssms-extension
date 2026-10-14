using Flank.Ssrs;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Asks for a stored procedure's parameter values, runs it to preview the
    /// results, and returns the values and results to build the output from.
    /// </summary>
    internal sealed class StoredProcedureRunDialog : Form
    {
        /// <summary>Rows kept per result set (shown in the grid and written to the output).</summary>
        public const int MaxRows = 50000;

        private readonly string procedureName;
        private readonly string outputName;
        private readonly IReadOnlyList<SprocParameter> parameters;
        private readonly IReadOnlyList<SprocParameter> outputParameters;
        private readonly Func<IDbConnection> openConnection;
        private readonly List<TextBox> inputs = new List<TextBox>();
        private readonly TabControl results;
        private readonly Label status;
        private readonly Button runButton;
        private readonly Button createButton;

        /// <param name="parameters">All of the procedure's parameters; OUTPUT ones get no input.</param>
        /// <param name="openConnection">Opens a connection to the procedure's database.</param>
        /// <param name="outputName">What's being created, for messages, e.g. "Excel (Power Query)".</param>
        public StoredProcedureRunDialog(
            string procedureName,
            IReadOnlyList<SprocParameter> parameters,
            Func<IDbConnection> openConnection,
            string createText,
            string outputName)
        {
            this.procedureName = procedureName;
            this.outputName = outputName;
            this.parameters = parameters.Where(p => !p.IsOutput).ToList();
            outputParameters = parameters.Where(p => p.IsOutput).ToList();
            parameters = this.parameters;
            this.openConnection = openConnection;

            Text = "Run " + procedureName;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            ClientSize = new Size(780, 600);
            MinimumSize = new Size(600, 450);

            var intro = new Label
            {
                Text = "Run the procedure to preview its results. The output starts with these " +
                       "values and rows; users can change the values and refresh.",
                Location = new Point(16, 12),
                Size = new Size(748, 34),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            };
            Controls.Add(intro);

            // Parameter inputs; scrolls if there are many.
            int inputsHeight = Math.Min(Math.Max(parameters.Count, 1) * 30, 240);
            var inputPanel = new Panel
            {
                Location = new Point(16, 50),
                Size = new Size(748, inputsHeight),
                AutoScroll = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            };
            Controls.Add(inputPanel);

            if (parameters.Count == 0)
            {
                inputPanel.Controls.Add(new Label
                {
                    Text = "This procedure has no parameters.",
                    Location = new Point(0, 6),
                    AutoSize = true,
                });
            }

            for (int i = 0; i < parameters.Count; i++)
            {
                var parameter = parameters[i];

                inputPanel.Controls.Add(new Label
                {
                    Text = $"{parameter.Name} ({parameter.SqlType})",
                    Location = new Point(0, i * 30 + 4),
                    Size = new Size(220, 20),
                    AutoEllipsis = true,
                });

                var input = new TextBox
                {
                    Location = new Point(225, i * 30),
                    Width = 300,
                    Name = "Parameter_" + parameter.Name.TrimStart('@'),
                };
                input.TextChanged += (s, e) => InvalidatePreview();
                inputs.Add(input);
                inputPanel.Controls.Add(input);
            }

            int y = 50 + inputsHeight + 8;

            if (parameters.Count > 0)
            {
                Controls.Add(new Label
                {
                    Text = "Leave a value blank to use the procedure's default.",
                    Location = new Point(16, y),
                    AutoSize = true,
                    ForeColor = SystemColors.GrayText,
                });
                y += 22;
            }

            runButton = new Button
            {
                Text = "Run",
                Location = new Point(16, y),
                Width = 90,
            };
            runButton.Click += async (s, e) => await RunAsync();
            Controls.Add(runButton);

            status = new Label
            {
                Location = new Point(116, y + 4),
                Size = new Size(648, 20),
                AutoEllipsis = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            };
            Controls.Add(status);

            y += 34;

            results = new TabControl
            {
                Location = new Point(16, y),
                Size = new Size(748, ClientSize.Height - y - 52),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            Controls.Add(results);

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
                if (!ConfirmFirstResultSetOnly())
                    return;

                DialogResult = DialogResult.OK;
                Close();
            };

            Controls.Add(cancelButton);
            Controls.Add(createButton);
            CancelButton = cancelButton;
            AcceptButton = runButton;
        }

        /// <summary>Typed values from the last run, in input-parameter order (null = default).</summary>
        public IReadOnlyList<object?> ParameterValues { get; private set; } = Array.Empty<object?>();

        /// <summary>Result sets from the last run.</summary>
        public IReadOnlyList<ResultSet> Results { get; private set; } = Array.Empty<ResultSet>();

        /// <summary>When the last run happened.</summary>
        public DateTime RanAt { get; private set; }

        /// <summary>
        /// Power Query and SSRS only ever read a procedure's first result set,
        /// so make sure the user knows the rest won't be in the output.
        /// </summary>
        private bool ConfirmFirstResultSetOnly()
        {
            if (Results.Count <= 1)
                return true;

            var answer = MessageBox.Show(
                this,
                $"{procedureName} returns {Results.Count} result sets.\n\n" +
                $"{outputName} only supports one result set per stored procedure, so only " +
                $"Result 1 ({Results[0].Columns.Count} columns) will be included. " +
                $"Results 2-{Results.Count} will be left out.\n\n" +
                "To include them, split the procedure into one procedure per result set.\n\n" +
                "Continue with Result 1 only?",
                "Flank: only the first result set is supported",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            return answer == DialogResult.OK;
        }

        private void InvalidatePreview()
        {
            if (createButton.Enabled)
            {
                createButton.Enabled = false;
                AcceptButton = runButton;
                status.ForeColor = SystemColors.ControlText;
                status.Text = "Values changed. Run again to preview.";
            }
        }

        private async Task RunAsync()
        {
            var values = parameters
                .Select((p, i) => StoredProcedureRunner.ParseValue(p.SqlType, inputs[i].Text))
                .ToList();

            var named = parameters
                .Select((p, i) => new KeyValuePair<string, object?>(p.Name, values[i]))
                .ToList();

            runButton.Enabled = false;
            createButton.Enabled = false;
            status.ForeColor = SystemColors.ControlText;
            status.Text = "Running...";
            UseWaitCursor = true;

            try
            {
                var runResults = await Task.Run(() =>
                {
                    using (var connection = openConnection())
                    {
                        return StoredProcedureRunner.Run(connection, procedureName, named, MaxRows, outputParameters);
                    }
                });

                if (runResults.Count == 0)
                {
                    status.ForeColor = Color.Firebrick;
                    status.Text = "The procedure didn't return a result set.";
                    return;
                }

                ShowResults(runResults);

                ParameterValues = values;
                Results = runResults;
                RanAt = DateTime.Now;
                createButton.Enabled = true;
                AcceptButton = createButton;

                status.Text = string.Join("; ", runResults.Select((r, i) =>
                    (runResults.Count > 1 ? $"Result {i + 1}: " : "") +
                    $"{r.Rows.Count:N0} row(s)" + (r.Truncated ? $" (first {MaxRows:N0} kept)" : "")));

                // Power Query and SSRS only ever read a procedure's first result set.
                if (runResults.Count > 1)
                    status.Text += ". Only Result 1 is used: Excel and SSRS read just the first result set.";
            }
            catch (Exception ex)
            {
                status.ForeColor = Color.Firebrick;
                status.Text = ex.Message;
            }
            finally
            {
                UseWaitCursor = false;
                runButton.Enabled = true;
            }
        }

        private void ShowResults(IReadOnlyList<ResultSet> runResults)
        {
            results.TabPages.Clear();

            for (int i = 0; i < runResults.Count; i++)
            {
                var result = runResults[i];
                var table = new DataTable();

                // DataTable needs unique column names; the grid just shows them.
                for (int c = 0; c < result.Columns.Count; c++)
                {
                    var column = table.Columns.Add("c" + c, typeof(object));
                    column.Caption = string.IsNullOrEmpty(result.Columns[c]) ? $"(column {c + 1})" : result.Columns[c];
                }

                foreach (var row in result.Rows)
                {
                    table.Rows.Add(row.Select(v => v ?? DBNull.Value).ToArray());
                }

                var grid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    RowHeadersVisible = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                    DataSource = table,
                };
                grid.DataBindingComplete += (s, e) =>
                {
                    foreach (DataGridViewColumn column in grid.Columns)
                        column.HeaderText = table.Columns[column.DataPropertyName].Caption;
                };

                var page = new TabPage(runResults.Count > 1 ? $"Result {i + 1}" : "Results");
                page.Controls.Add(grid);
                results.TabPages.Add(page);
            }
        }
    }
}
