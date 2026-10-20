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
    /// Parameter inputs, a Run button, and the results of the run, one tab per
    /// result set. The host supplies how to run (as whom, against what).
    /// </summary>
    internal sealed class PreviewPanel : UserControl
    {
        /// <summary>Rows kept per result set.</summary>
        public const int MaxRows = 50000;

        private readonly IReadOnlyList<SprocParameter> parameters;
        private readonly IReadOnlyList<ParameterSettings?> settings;
        // A TextBox, or a ComboBox of Choices for a parameter with a dropdown list.
        private readonly List<Control> inputs = new List<Control>();

        /// <summary>One of a list parameter's choices: shown by its label, passing its value.</summary>
        private sealed class Choice
        {
            public static readonly Choice Blank = new Choice(null, "(blank)");

            public Choice(object? value, object? label)
            {
                Value = value;
                Label = Convert.ToString(label, System.Globalization.CultureInfo.CurrentCulture) ?? "";
            }

            public object? Value { get; }

            public string Label { get; }

            public bool Matches(string? text) =>
                text != null && Value != null &&
                string.Equals(Convert.ToString(Value, System.Globalization.CultureInfo.InvariantCulture), text.Trim(), StringComparison.OrdinalIgnoreCase);

            public override string ToString() => Label;
        }

        /// <summary>The value an input holds: typed for a box, the chosen value for a list; null if blank.</summary>
        private object? ValueOf(int index)
        {
            if (inputs[index] is ComboBox choices)
                return (choices.SelectedItem as Choice)?.Value;

            return StoredProcedureRunner.ParseValue(parameters[index].SqlType, inputs[index].Text);
        }
        private readonly TabControl results;
        private readonly Label status;

        /// <param name="parameters">Input parameters; each gets a box.</param>
        /// <param name="size">Initial size; controls are laid out to it and anchored.</param>
        /// <param name="reviewed">
        /// From the parameter review page, if there was one: labels, starting
        /// values and what a blank value means.
        /// </param>
        public PreviewPanel(IReadOnlyList<SprocParameter> parameters, Size size, IReadOnlyList<ParameterSettings>? reviewed = null)
        {
            this.parameters = parameters;
            settings = parameters
                .Select(p => reviewed?.FirstOrDefault(s => string.Equals(s.SqlName, p.Name, StringComparison.OrdinalIgnoreCase)))
                .ToList();
            Size = size;

            int inputsHeight = parameters.Count == 0 ? 24 : Math.Min(parameters.Count * 30, 240);
            var inputPanel = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(Width, inputsHeight),
                AutoScroll = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            };
            Controls.Add(inputPanel);

            if (parameters.Count == 0)
            {
                inputPanel.Controls.Add(new Label { Text = "No parameters.", Location = new Point(0, 4), AutoSize = true });
            }

            for (int i = 0; i < parameters.Count; i++)
            {
                var parameter = parameters[i];
                var setting = settings[i];

                inputPanel.Controls.Add(new Label
                {
                    Text = setting != null
                        ? $"{setting.DisplayName} ({parameter.SqlType}){(setting.IfBlank == BlankValue.Required ? " *" : "")}"
                        : $"{parameter.Name} ({parameter.SqlType})",
                    Location = new Point(0, i * 30 + 4),
                    Size = new Size(220, 20),
                    AutoEllipsis = true,
                });

                Control input;

                if (setting?.HasList == true)
                {
                    // The same choices people get in the report or workbook.
                    var choices = new ComboBox
                    {
                        Location = new Point(225, i * 30),
                        Width = 300,
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        Name = "Parameter_" + parameter.Name.TrimStart('@'),
                    };

                    if (setting.IfBlank != BlankValue.Required)
                        choices.Items.Add(Choice.Blank);

                    foreach (var row in setting.ListRows)
                        choices.Items.Add(new Choice(row.Key, row.Value));

                    choices.SelectedItem =
                        choices.Items.OfType<Choice>().FirstOrDefault(c => c.Matches(setting.DefaultValue)) ??
                        (setting.IfBlank != BlankValue.Required ? Choice.Blank : null);

                    choices.SelectedIndexChanged += (s, e) => Invalidate("Values changed. Run again to preview.");
                    input = choices;
                }
                else
                {
                    input = new TextBox
                    {
                        Text = setting?.DefaultValue ?? "",
                        Location = new Point(225, i * 30),
                        Width = 300,
                        Name = "Parameter_" + parameter.Name.TrimStart('@'),
                    };
                    input.TextChanged += (s, e) => Invalidate("Values changed. Run again to preview.");
                }

                inputs.Add(input);
                inputPanel.Controls.Add(input);
            }

            int y = inputsHeight + 6;

            if (parameters.Count > 0)
            {
                Controls.Add(new Label
                {
                    Text = settings.Any(s => s != null)
                        ? "* Required. Other blank values are left out (the procedure's default) or NULL, as set on the review page."
                        : "Leave a value blank to use the procedure's default.",
                    Location = new Point(0, y),
                    AutoSize = true,
                    ForeColor = SystemColors.GrayText,
                });
                y += 22;
            }

            // Grows to fit e.g. "Run as DOMAIN\someone"; the status follows it.
            RunButton = new Button
            {
                Text = "Run",
                Location = new Point(0, y),
                MinimumSize = new Size(140, 0),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
            };
            RunButton.Click += async (s, e) => await RunAsync();
            Controls.Add(RunButton);

            status = new Label
            {
                Location = new Point(150, y + 4),
                Size = new Size(Width - 150, 34),
                AutoEllipsis = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            };
            Controls.Add(status);
            RunButton.SizeChanged += (s, e) =>
            {
                status.Left = RunButton.Right + 10;
                status.Width = Width - status.Left;
            };

            y += 40;

            results = new TabControl
            {
                Location = new Point(0, y),
                Size = new Size(Width, Math.Max(Height - y, 100)),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            Controls.Add(results);
        }

        /// <summary>Runs with the typed values (in parameter order); called off the UI thread.</summary>
        public Func<IReadOnlyList<object?>, IReadOnlyList<ResultSet>>? Runner { get; set; }

        public Button RunButton { get; }

        /// <summary>Puts the cursor in the first value box, or on Run when there are none.</summary>
        public void FocusFirstInput()
        {
            if (inputs.Count > 0)
                inputs[0].Select();
            else
                RunButton.Select();
        }

        public IReadOnlyList<object?> ParameterValues { get; private set; } = Array.Empty<object?>();

        public IReadOnlyList<ResultSet> Results { get; private set; } = Array.Empty<ResultSet>();

        public DateTime RanAt { get; private set; }

        /// <summary>A successful run whose values are still current.</summary>
        public bool HasResults => Results.Count > 0;

        /// <summary>Raised when <see cref="HasResults"/> may have changed.</summary>
        public event EventHandler? ResultsChanged;

        /// <summary>Raised when a run fails, with the error.</summary>
        public event EventHandler<Exception>? RunFailed;

        public void SetStatus(string text, Color? color = null)
        {
            status.ForeColor = color ?? SystemColors.ControlText;
            status.Text = text;
        }

        private void Invalidate(string message)
        {
            if (!HasResults)
                return;

            Results = Array.Empty<ResultSet>();
            SetStatus(message);
            ResultsChanged?.Invoke(this, EventArgs.Empty);
        }

        public async Task RunAsync()
        {
            if (Runner == null)
                return;

            for (int i = 0; i < parameters.Count; i++)
            {
                if (settings[i]?.IfBlank == BlankValue.Required && ValueOf(i) == null)
                {
                    SetStatus($"Enter a value for {settings[i]!.DisplayName}: it's required.", Color.Firebrick);
                    inputs[i].Select();
                    return;
                }
            }

            // Null leaves a parameter out (the procedure's default); DBNull passes NULL.
            var values = parameters
                .Select((p, i) =>
                    ValueOf(i) ??
                    (settings[i]?.IfBlank == BlankValue.PassNull ? DBNull.Value : null))
                .ToList();

            var runner = Runner;
            RunButton.Enabled = false;
            Results = Array.Empty<ResultSet>();
            ResultsChanged?.Invoke(this, EventArgs.Empty);
            SetStatus("Running...");
            UseWaitCursor = true;

            try
            {
                var runResults = await Task.Run(() => runner(values));

                if (runResults.Count == 0)
                {
                    SetStatus("No result set came back.", Color.Firebrick);
                    return;
                }

                ShowResults(runResults);
                ParameterValues = values;
                Results = runResults;
                RanAt = DateTime.Now;

                string summary = string.Join("; ", runResults.Select((r, i) =>
                    (runResults.Count > 1 ? $"Result {i + 1}: " : "") +
                    (r.Truncated ? $"over {MaxRows:N0} rows" : $"{r.Rows.Count:N0} row(s)")));

                if (runResults[0].Truncated)
                    summary += $". Showing the first {MaxRows:N0}.";

                SetStatus(summary);
                ResultsChanged?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                SetStatus(ex.Message, Color.Firebrick);
                RunFailed?.Invoke(this, ex);
            }
            finally
            {
                UseWaitCursor = false;
                RunButton.Enabled = true;
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
                    table.Rows.Add(row.Select(v => v ?? DBNull.Value).ToArray());

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

        /// <summary>
        /// Power Query and SSRS only ever read the first result set: make sure
        /// the user knows the rest won't be in the output.
        /// </summary>
        public static bool ConfirmFirstResultSetOnly(IWin32Window owner, string what, IReadOnlyList<ResultSet> results, string outputName)
        {
            if (results.Count <= 1)
                return true;

            var answer = MessageBox.Show(
                owner,
                $"{what} returns {results.Count} result sets.\n\n" +
                $"{outputName} only supports one result set, so only Result 1 " +
                $"({results[0].Columns.Count} columns) will be included. " +
                $"Results 2-{results.Count} will be left out.\n\n" +
                "To include them, split the procedure into one procedure per result set.\n\n" +
                "Continue with Result 1 only?",
                "Flank: only the first result set is supported",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            return answer == DialogResult.OK;
        }
    }
}
