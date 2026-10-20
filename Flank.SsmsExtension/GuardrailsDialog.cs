using Flank.Ssrs;
using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// Guardrails for one parameter, beyond its type: for now, a dropdown of
    /// choices from a stored procedure, with the column shown and the column
    /// passed chosen by the user. Laid out in sections so others (ranges,
    /// patterns) can follow.
    /// </summary>
    internal sealed class GuardrailsDialog : Form
    {
        private readonly ParameterSettings setting;
        private readonly Func<IDbConnection> openConnection;
        private readonly CheckBox useList;
        private readonly ComboBox procedureComboBox;
        private readonly Label listStatus;
        private readonly ComboBox labelComboBox;
        private readonly ComboBox valueComboBox;
        private readonly DataGridView preview;
        private ProcedureList.Description? list;

        public GuardrailsDialog(ParameterSettings setting, Func<IDbConnection> openConnection)
        {
            this.setting = setting;
            this.openConnection = openConnection;

            Text = "Guardrails";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(560, 470);

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
                Text = "Choose from a list of values returned by a stored procedure",
                Location = new Point(18, 102),
                Width = 520,
                Checked = setting.HasList,
            };
            useList.CheckedChanged += (s, e) => UpdateEnabled();

            Controls.Add(new Label { Text = "Procedure:", Location = new Point(36, 135), Width = 110 });

            // The database's procedures without parameters; others (all parameters
            // defaulted) can be typed. Picking or leaving the box loads the columns.
            procedureComboBox = new ComboBox
            {
                Text = setting.ListProcedure ?? "",
                Location = new Point(150, 132),
                Width = 388,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems,
            };
            procedureComboBox.SelectionChangeCommitted += (s, e) =>
                BeginInvoke((Action)(() => LoadList(keepSelection: false)));
            procedureComboBox.TextUpdate += (s, e) =>
            {
                list = null;
                ShowColumns();
            };
            procedureComboBox.Leave += (s, e) =>
            {
                if (list == null && procedureComboBox.Text.Trim().Length > 0)
                    LoadList(keepSelection: false);
            };

            listStatus = new Label
            {
                Text = "Pick the procedure that returns the choices. Procedures that take parameters aren't listed; " +
                       "type one whose parameters all have defaults.",
                ForeColor = SystemColors.GrayText,
                Location = new Point(150, 158),
                Size = new Size(388, 32),
            };

            Controls.Add(new Label { Text = "Show:", Location = new Point(36, 197), Width = 110 });
            labelComboBox = new ComboBox { Location = new Point(150, 194), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
            Controls.Add(new Label { Text = "what people pick from", ForeColor = SystemColors.GrayText, Location = new Point(358, 197), AutoSize = true });

            Controls.Add(new Label { Text = "Pass:", Location = new Point(36, 229), Width = 110 });
            valueComboBox = new ComboBox { Location = new Point(150, 226), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
            Controls.Add(new Label { Text = "what the procedure receives", ForeColor = SystemColors.GrayText, Location = new Point(358, 229), AutoSize = true });

            labelComboBox.SelectedIndexChanged += (s, e) => ShowPreview();
            valueComboBox.SelectedIndexChanged += (s, e) => ShowPreview();

            preview = new DataGridView
            {
                Location = new Point(36, 262),
                Size = new Size(502, 150),
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = SystemColors.Window,
            };

            var okButton = new Button { Text = "OK", Location = new Point(368, 428), Width = 80 };
            okButton.Click += (s, e) => Save();

            var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(458, 428), Width = 80 };

            Controls.AddRange(new Control[]
            {
                useList, procedureComboBox, listStatus, labelComboBox, valueComboBox, preview, okButton, cancelButton,
            });

            AcceptButton = okButton;
            CancelButton = cancelButton;

            UpdateEnabled();

            Shown += (s, e) =>
            {
                LoadProcedures();

                // Editing an existing list: load it so its columns can be changed.
                if (setting.HasList)
                    LoadList(keepSelection: true);
            };
        }

        private void LoadProcedures()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                string typed = procedureComboBox.Text;

                using (var connection = openConnection())
                {
                    procedureComboBox.Items.AddRange(ProcedureList.Candidates(connection).Cast<object>().ToArray());
                }

                procedureComboBox.Text = typed;
            }
            catch (Exception ex)
            {
                // The box still takes a typed name.
                System.Diagnostics.Debug.WriteLine($"Flank: could not list procedures: {ex.Message}");
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void UpdateEnabled()
        {
            bool on = useList.Checked;
            procedureComboBox.Enabled = on;
            labelComboBox.Enabled = on && list != null;
            valueComboBox.Enabled = on && list != null;
            preview.Enabled = on;
        }

        private void LoadList(bool keepSelection)
        {
            string name = procedureComboBox.Text.Trim();

            if (name.Length == 0)
            {
                SetStatus("Enter a stored procedure, e.g. dbo.get_leagues.", error: true);
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                using (var connection = openConnection())
                {
                    list = ProcedureList.Describe(name, connection);
                }

                // Text, not TextUpdate (which only typing raises): the list stays loaded.
                procedureComboBox.Text = list.Procedure;
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

        /// <summary>The first rows, Show then Pass, as people will see and the procedure receive them.</summary>
        private void ShowPreview()
        {
            preview.Columns.Clear();
            preview.Rows.Clear();

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

            if (list == null)
            {
                LoadList(keepSelection: setting.HasList);

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
        public static string Summary(ParameterSettings setting) =>
            setting.HasList
                ? $"Dropdown: {setting.ListProcedure} ({setting.ListLabelColumn} → {setting.ListValueColumn})"
                : "";
    }
}
