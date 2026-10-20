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
    /// Guardrails for one parameter, beyond its type: for now, a dropdown of
    /// choices from a stored procedure, a query, or values typed in, with the
    /// column shown and the column passed chosen by the user. Laid out in
    /// sections so others (ranges, patterns) can follow.
    /// </summary>
    internal sealed class GuardrailsDialog : Form
    {
        private const int Indent = 36;
        private const int PanelWidth = 546;

        private readonly ParameterSettings setting;
        private readonly Func<IDbConnection> openConnection;
        private readonly CheckBox useList;

        private readonly RadioButton fromProcedure;
        private readonly RadioButton fromQuery;
        private readonly RadioButton fromValues;

        private readonly Panel procedurePanel;
        private readonly ComboBox procedureComboBox;

        private readonly Panel queryPanel;
        private readonly TextBox queryTextBox;

        private readonly Panel valuesPanel;
        private readonly RadioButton valuesOnly;
        private readonly RadioButton valuesAndLabels;
        private readonly TextBox valuesTextBox;
        private readonly Label labelsLabel;
        private readonly TextBox labelsTextBox;

        private readonly Label listStatus;
        private readonly Label showLabel;
        private readonly Label passLabel;
        private readonly Label showHint;
        private readonly Label passHint;
        private readonly ComboBox labelComboBox;
        private readonly ComboBox valueComboBox;
        private readonly DataGridView preview;

        // From a run of the procedure or query.
        private ProcedureList.Description? list;

        // Typed-in values, when they parse.
        private IReadOnlyList<KeyValuePair<object?, object?>>? typedRows;

        public GuardrailsDialog(ParameterSettings setting, Func<IDbConnection> openConnection)
        {
            this.setting = setting;
            this.openConnection = openConnection;

            Text = "Guardrails";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(600, 650);

            Controls.Add(new Label
            {
                Text = "Guardrails for " + setting.DisplayName,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 13, FontStyle.Bold),
                Location = new Point(16, 14),
                AutoSize = true,
            });

            Controls.Add(new Label
            {
                Text = $"{setting.SqlName} {ProcedureCall.TypeDeclaration(setting.Parameter)}",
                ForeColor = SystemColors.GrayText,
                Location = new Point(18, 44),
                AutoSize = true,
            });

            // ---- Dropdown list
            Controls.Add(new Label
            {
                Text = "Dropdown list",
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                Location = new Point(16, 80),
                AutoSize = true,
            });

            useList = new CheckBox
            {
                Text = "People choose from a list",
                Location = new Point(18, 102),
                Width = 520,
                Checked = setting.HasList,
            };
            useList.CheckedChanged += (s, e) => UpdateEnabled();

            Controls.Add(new Label { Text = "Choices from:", Location = new Point(Indent, 133), AutoSize = true });

            fromProcedure = new RadioButton { Text = "A stored procedure", Location = new Point(130, 130), Width = 140 };
            fromQuery = new RadioButton { Text = "A query", Location = new Point(275, 130), Width = 90 };
            fromValues = new RadioButton { Text = "Values typed here", Location = new Point(370, 130), Width = 160 };

            // ---- From a procedure
            procedurePanel = new Panel { Location = new Point(Indent, 162), Size = new Size(PanelWidth, 190) };
            procedurePanel.Controls.Add(new Label { Text = "Procedure:", Location = new Point(0, 3), Width = 90 });

            // A plain list: typing to match only works from the start of the
            // name, which is always "[schema]".
            procedureComboBox = new ComboBox
            {
                Location = new Point(94, 0),
                Width = PanelWidth - 94,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Name = "ListProcedure",
            };
            procedureComboBox.SelectionChangeCommitted += (s, e) =>
                BeginInvoke((Action)(() =>
                {
                    // Picking one means wanting the dropdown: tick it too.
                    useList.Checked = true;
                    LoadList(keepSelection: false);
                }));

            procedurePanel.Controls.Add(procedureComboBox);
            procedurePanel.Controls.Add(Hint(
                "Procedures that take no parameters. For one that does, choose A query and call it with values, " +
                "e.g. EXEC dbo.get_teams @active = 1.",
                new Point(94, 28),
                PanelWidth - 94));

            // ---- From a query
            queryPanel = new Panel { Location = new Point(Indent, 162), Size = new Size(PanelWidth, 190) };

            queryTextBox = new TextBox
            {
                Location = new Point(0, 0),
                Size = new Size(PanelWidth, 120),
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                AcceptsReturn = true,
                AcceptsTab = true,
                Font = new Font("Consolas", 9),
                Text = setting.ListQuery ?? "",
                Name = "ListQuery",
            };
            queryTextBox.TextChanged += (s, e) =>
            {
                if (queryTextBox.Focused)
                    useList.Checked = true;

                list = null;
                ShowColumns();
            };

            var runQuery = new Button { Text = "Run query", Location = new Point(0, 128), Width = 100 };
            runQuery.Click += (s, e) =>
            {
                useList.Checked = true;
                LoadList(keepSelection: false);
            };

            queryPanel.Controls.Add(queryTextBox);
            queryPanel.Controls.Add(runQuery);
            queryPanel.Controls.Add(Hint(
                "Any query that returns the choices, e.g. SELECT id, name FROM dbo.leagues WHERE active = 1. " +
                "Run it, then choose its columns.",
                new Point(108, 128),
                PanelWidth - 108));

            // ---- Typed-in values
            valuesPanel = new Panel { Location = new Point(Indent, 162), Size = new Size(PanelWidth, 190) };

            bool labelled = ProcedureList.HasLabels(setting);
            valuesOnly = new RadioButton { Text = "Values only", Location = new Point(0, 0), Width = 110, Checked = !labelled };
            valuesAndLabels = new RadioButton { Text = "Values, each with a label people see", Location = new Point(115, 0), Width = 300, Checked = labelled };

            valuesPanel.Controls.Add(new Label { Text = "Values (one per line):", Location = new Point(0, 28), AutoSize = true });

            valuesTextBox = ValuesBox(new Point(0, 46), "ListValues");
            labelsLabel = new Label { Text = "Labels (same lines):", Location = new Point(PanelWidth / 2 + 6, 28), AutoSize = true };
            labelsTextBox = ValuesBox(new Point(PanelWidth / 2 + 6, 46), "ListLabels");

            if (setting.ListSource == ListSource.Values)
            {
                valuesTextBox.Text = string.Join(Environment.NewLine, setting.ListRows.Select(r => ProcedureList.FormatValue(r.Key)));
                labelsTextBox.Text = labelled
                    ? string.Join(Environment.NewLine, setting.ListRows.Select(r => Convert.ToString(r.Value)))
                    : "";
            }

            valuesOnly.CheckedChanged += (s, e) => ParseValues();
            valuesTextBox.TextChanged += (s, e) =>
            {
                if (valuesTextBox.Focused)
                    useList.Checked = true;

                ParseValues();
            };
            labelsTextBox.TextChanged += (s, e) => ParseValues();

            valuesPanel.Controls.AddRange(new Control[] { valuesOnly, valuesAndLabels, valuesTextBox, labelsLabel, labelsTextBox });

            // ---- Status, columns, preview (shared)
            listStatus = new Label
            {
                ForeColor = SystemColors.GrayText,
                Location = new Point(Indent, 356),
                Size = new Size(PanelWidth, 32),
            };

            showLabel = new Label { Text = "Show:", Location = new Point(Indent, 395), Width = 60 };
            labelComboBox = new ComboBox { Location = new Point(100, 392), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
            showHint = Hint("what people pick from", new Point(308, 395), 270);

            passLabel = new Label { Text = "Pass:", Location = new Point(Indent, 427), Width = 60 };
            valueComboBox = new ComboBox { Location = new Point(100, 424), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
            passHint = Hint("what the parameter receives", new Point(308, 427), 270);

            labelComboBox.SelectedIndexChanged += (s, e) => ShowPreview();
            valueComboBox.SelectedIndexChanged += (s, e) => ShowPreview();

            preview = new DataGridView
            {
                Location = new Point(Indent, 458),
                Size = new Size(PanelWidth, 140),
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = SystemColors.Window,
            };

            var okButton = new Button { Text = "OK", Location = new Point(410, 610), Width = 80 };
            okButton.Click += (s, e) => Save();

            var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(502, 610), Width = 80 };

            Controls.AddRange(new Control[]
            {
                useList, fromProcedure, fromQuery, fromValues, procedurePanel, queryPanel, valuesPanel,
                listStatus, showLabel, labelComboBox, showHint, passLabel, valueComboBox, passHint,
                preview, okButton, cancelButton,
            });

            CancelButton = cancelButton;

            fromProcedure.Checked = !setting.HasList || setting.ListSource == ListSource.Procedure;
            fromQuery.Checked = setting.HasList && setting.ListSource == ListSource.Query;
            fromValues.Checked = setting.HasList && setting.ListSource == ListSource.Values;

            foreach (var radio in new[] { fromProcedure, fromQuery, fromValues })
                radio.CheckedChanged += (s, e) =>
                {
                    if (((RadioButton)s).Checked)
                    {
                        useList.Checked = true;
                        SourceChanged();
                    }
                };

            SourceChanged();

            Shown += (s, e) =>
            {
                LoadProcedures();

                // Editing an existing list: load it so its columns can be changed.
                if (setting.HasList && setting.ListSource != ListSource.Values)
                    LoadList(keepSelection: true);
            };
        }

        private ListSource Source =>
            fromValues.Checked ? ListSource.Values : fromQuery.Checked ? ListSource.Query : ListSource.Procedure;

        private static Label Hint(string text, Point location, int width) => new Label
        {
            Text = text,
            ForeColor = SystemColors.GrayText,
            Location = location,
            Size = new Size(width, 32),
        };

        private static TextBox ValuesBox(Point location, string name) => new TextBox
        {
            Location = location,
            Size = new Size(PanelWidth / 2 - 6, 140),
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            AcceptsReturn = true,
            Name = name,
        };

        /// <summary>A procedure in the list: [schema].[name], shown as schema.name.</summary>
        private sealed class ProcedureItem
        {
            public ProcedureItem(string name) => Name = name;

            public string Name { get; }

            public override string ToString() => Name.Replace("[", "").Replace("]", "");
        }

        private void LoadProcedures()
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                using (var connection = openConnection())
                {
                    procedureComboBox.Items.AddRange(
                        ProcedureList.Candidates(connection).Select(n => (object)new ProcedureItem(n)).ToArray());
                }

                // The procedure already chosen, even if it now takes parameters.
                if (setting.ListSource == ListSource.Procedure && setting.ListProcedure != null)
                {
                    var current = procedureComboBox.Items.OfType<ProcedureItem>()
                        .FirstOrDefault(i => string.Equals(i.Name, setting.ListProcedure, StringComparison.OrdinalIgnoreCase));

                    if (current == null)
                    {
                        current = new ProcedureItem(setting.ListProcedure);
                        procedureComboBox.Items.Insert(0, current);
                    }

                    procedureComboBox.SelectedItem = current;
                }
            }
            catch (Exception ex)
            {
                SetStatus("Flank couldn't list the database's procedures: " + ex.Message, error: true);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void SourceChanged()
        {
            var source = Source;

            procedurePanel.Visible = source == ListSource.Procedure;
            queryPanel.Visible = source == ListSource.Query;
            valuesPanel.Visible = source == ListSource.Values;

            bool columns = source != ListSource.Values;
            foreach (Control control in new Control[] { showLabel, labelComboBox, showHint, passLabel, valueComboBox, passHint })
                control.Visible = columns;

            preview.Top = columns ? 458 : 392;
            preview.Height = columns ? 140 : 206;

            list = null;
            typedRows = null;

            if (source == ListSource.Values)
            {
                ParseValues();
            }
            else
            {
                SetStatus(
                    source == ListSource.Procedure
                        ? "Pick the procedure that returns the choices."
                        : "Write the query, then Run query.",
                    error: false);

                // Back to a procedure that was already picked: show it again.
                if (source == ListSource.Procedure && procedureComboBox.SelectedItem != null && IsHandleCreated)
                    LoadList(keepSelection: setting.ListSource == ListSource.Procedure);
                else
                    ShowColumns();
            }

            UpdateEnabled();
        }

        private void UpdateEnabled()
        {
            // The list's controls stay usable before the box is ticked: using them ticks it.
            labelComboBox.Enabled = list != null;
            valueComboBox.Enabled = list != null;
        }

        private void LoadList(bool keepSelection)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                if (Source == ListSource.Procedure)
                {
                    if (!(procedureComboBox.SelectedItem is ProcedureItem item))
                    {
                        SetStatus("Pick the procedure that returns the choices.", error: true);
                        return;
                    }

                    using (var connection = openConnection())
                        list = ProcedureList.Describe(item.Name, connection);
                }
                else
                {
                    using (var connection = openConnection())
                        list = ProcedureList.DescribeQuery(queryTextBox.Text, connection);
                }

                SetStatus($"{list.Rows.Count:N0} choice(s), {list.Columns.Count} column(s).", error: false);
            }
            catch (Exception ex)
            {
                list = null;
                SetStatus(ex.Message, error: true);
            }
            finally
            {
                Cursor = Cursors.Default;
            }

            ShowColumns(keepSelection);
        }

        /// <summary>Checks the typed-in values as they're typed, and previews them.</summary>
        private void ParseValues()
        {
            labelsLabel.Visible = valuesAndLabels.Checked;
            labelsTextBox.Visible = valuesAndLabels.Checked;
            valuesTextBox.Width = valuesAndLabels.Checked ? PanelWidth / 2 - 6 : PanelWidth;

            typedRows = ProcedureList.ParseValues(
                setting.Parameter,
                valuesTextBox.Text,
                valuesAndLabels.Checked ? labelsTextBox.Text : null,
                out string? problem);

            if (typedRows != null)
                SetStatus($"{typedRows.Count:N0} choice(s).", error: false);
            else if (valuesTextBox.Text.Trim().Length == 0)
                SetStatus($"Type the values, one per line, as {ProcedureCall.TypeDeclaration(setting.Parameter)}.", error: false);
            else
                SetStatus(problem!, error: true);

            ShowPreview();
        }

        private void SetStatus(string text, bool error)
        {
            listStatus.Text = text;
            listStatus.ForeColor = error ? Color.Firebrick : SystemColors.GrayText;
        }

        private void ShowColumns(bool keepSelection = false)
        {
            labelComboBox.Items.Clear();
            valueComboBox.Items.Clear();

            if (list != null)
            {
                labelComboBox.Items.AddRange(list.Columns.Cast<object>().ToArray());
                valueComboBox.Items.AddRange(list.Columns.Cast<object>().ToArray());

                string label = keepSelection && setting.ListLabelColumn != null ? setting.ListLabelColumn : GuessLabel(list);
                string value = keepSelection && setting.ListValueColumn != null ? setting.ListValueColumn : GuessValue(list, label);

                labelComboBox.SelectedItem = list.Columns.FirstOrDefault(c => string.Equals(c, label, StringComparison.OrdinalIgnoreCase));
                valueComboBox.SelectedItem = list.Columns.FirstOrDefault(c => string.Equals(c, value, StringComparison.OrdinalIgnoreCase));
            }

            ShowPreview();
            UpdateEnabled();
        }

        /// <summary>The first column holding text (a name), or else the first column.</summary>
        private static string GuessLabel(ProcedureList.Description list)
        {
            for (int c = 0; c < list.Columns.Count; c++)
            {
                if (list.Rows.Any(row => row[c] is string))
                    return list.Columns[c];
            }

            return list.Columns[0];
        }

        /// <summary>An id-like column other than the label, or else the first other column, or the label itself.</summary>
        private static string GuessValue(ProcedureList.Description list, string label)
        {
            var others = list.Columns.Where(c => !string.Equals(c, label, StringComparison.OrdinalIgnoreCase)).ToList();

            return others.FirstOrDefault(c => c.EndsWith("id", StringComparison.OrdinalIgnoreCase))
                ?? others.FirstOrDefault()
                ?? label;
        }

        /// <summary>The first choices, Show then Pass, as people will see and the parameter receive them.</summary>
        private void ShowPreview()
        {
            preview.Columns.Clear();
            preview.Rows.Clear();

            if (Source == ListSource.Values)
            {
                if (typedRows == null)
                    return;

                preview.Columns.Add("Show", "Show");
                preview.Columns.Add("Pass", "Pass");

                foreach (var row in typedRows.Take(200))
                    preview.Rows.Add(row.Value, ProcedureList.FormatValue(row.Key));

                return;
            }

            if (list == null || labelComboBox.SelectedItem == null || valueComboBox.SelectedItem == null)
                return;

            int label = list.Columns.ToList().IndexOf((string)labelComboBox.SelectedItem);
            int value = list.Columns.ToList().IndexOf((string)valueComboBox.SelectedItem);

            preview.Columns.Add("Show", "Show: " + list.Columns[label]);
            preview.Columns.Add("Pass", "Pass: " + list.Columns[value]);

            foreach (var row in list.Rows.Take(50))
                preview.Rows.Add(row[label], row[value]);
        }

        private void Save()
        {
            if (!useList.Checked)
            {
                ProcedureList.Clear(setting);
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            if (Source == ListSource.Values)
            {
                ParseValues();

                if (typedRows == null)
                    return;

                ProcedureList.ApplyValues(setting, typedRows, valuesAndLabels.Checked);
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            if (list == null)
            {
                LoadList(keepSelection: setting.HasList && setting.ListSource == Source);

                if (list == null)
                    return;
            }

            if (labelComboBox.SelectedItem == null || valueComboBox.SelectedItem == null)
            {
                SetStatus("Choose the column to show and the column to pass.", error: true);
                return;
            }

            ProcedureList.Apply(setting, list, (string)valueComboBox.SelectedItem, (string)labelComboBox.SelectedItem);
            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>A one-line description of a parameter's guardrails, for the review page.</summary>
        public static string Summary(ParameterSettings setting)
        {
            if (!setting.HasList)
                return "";

            switch (setting.ListSource)
            {
                case ListSource.Values:
                    return $"Dropdown: {setting.ListRows.Count} value(s)" + (ProcedureList.HasLabels(setting) ? " with labels" : "");

                case ListSource.Query:
                    return $"Dropdown: query ({setting.ListLabelColumn} → {setting.ListValueColumn})";

                default:
                    return $"Dropdown: {setting.ListProcedure} ({setting.ListLabelColumn} → {setting.ListValueColumn})";
            }
        }
    }
}
