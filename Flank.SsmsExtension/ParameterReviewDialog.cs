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
    /// Reviews a stored procedure's input parameters before creating a report
    /// or workbook from it: what people see each one as, a note for them, the
    /// value it starts with, and what leaving it blank means.
    /// </summary>
    internal sealed class ParameterReviewDialog : Form
    {
        private const string RequiredText = "Required";
        private const string NullText = "Pass NULL";

        private readonly IReadOnlyList<ParameterSettings> settings;
        private readonly Func<IDbConnection> openConnection;
        private readonly DataGridView grid;

        /// <param name="outputName">What's being created, e.g. "Excel workbook".</param>
        /// <param name="openConnection">Opens a connection to the database, to check list procedures.</param>
        public ParameterReviewDialog(
            string procedureName,
            IReadOnlyList<ParameterSettings> settings,
            string outputName,
            Func<IDbConnection> openConnection)
        {
            this.settings = settings;
            this.openConnection = openConnection;

            Text = "Review parameters: " + procedureName;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            ClientSize = new Size(900, 120 + Math.Min(settings.Count, 10) * 26 + 80);
            MinimumSize = new Size(700, 300);

            Controls.Add(new Label
            {
                Text =
                    $"How people using the {outputName} see and fill in each parameter. Defaults come from the " +
                    "procedure where it declares them. \"Use procedure default\" leaves the parameter out of the call " +
                    "when it's blank, so the procedure's own default applies. \"Choices from\" makes it a dropdown: " +
                    "name a procedure whose first column is the value and second (optional) the label.",
                Location = new Point(16, 12),
                Size = new Size(868, 48),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            });

            grid = new DataGridView
            {
                Location = new Point(16, 64),
                Size = new Size(868, ClientSize.Height - 64 - 56),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                EditMode = DataGridViewEditMode.EditOnEnter,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            };

            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Parameter", HeaderText = "Parameter", ReadOnly = true, FillWeight = 22 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "DisplayName", HeaderText = "Shown as", FillWeight = 18 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Note", HeaderText = "Note for the user", FillWeight = 24 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Default", HeaderText = "Starts with", FillWeight = 12 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "List", HeaderText = "Choices from (procedure)", FillWeight = 20 });
            grid.Columns.Add(new DataGridViewComboBoxColumn
            {
                Name = "IfBlank",
                HeaderText = "If left blank",
                FillWeight = 20,
                FlatStyle = FlatStyle.Flat,
            });

            grid.Columns["Parameter"].DefaultCellStyle.BackColor = SystemColors.Control;

            foreach (var setting in settings)
            {
                var parameter = setting.Parameter;
                int row = grid.Rows.Add(
                    $"{parameter.Name} {ProcedureCall.TypeDeclaration(parameter)}",
                    setting.DisplayName,
                    setting.Note,
                    setting.DefaultValue ?? "",
                    setting.ListProcedure ?? "",
                    null);

                var choice = (DataGridViewComboBoxCell)grid.Rows[row].Cells["IfBlank"];
                choice.Items.Add(RequiredText);
                if (parameter.HasDefault)
                    choice.Items.Add(UseDefaultText(parameter));
                choice.Items.Add(NullText);
                choice.Value = setting.IfBlank switch
                {
                    BlankValue.UseProcedureDefault when parameter.HasDefault => UseDefaultText(parameter),
                    BlankValue.PassNull => NullText,
                    _ => RequiredText,
                };
            }

            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(ClientSize.Width - 106, ClientSize.Height - 40),
                Width = 90,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };

            var nextButton = new Button
            {
                Text = "Next",
                Location = new Point(ClientSize.Width - 206, ClientSize.Height - 40),
                Width = 90,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            };
            nextButton.Click += (s, e) => Save();

            Controls.Add(grid);
            Controls.Add(cancelButton);
            Controls.Add(nextButton);
            AcceptButton = nextButton;
            CancelButton = cancelButton;
        }

        private static string UseDefaultText(SprocParameter parameter) =>
            parameter.DefaultValue != null
                ? $"Use procedure default ({parameter.DefaultValue})"
                : "Use procedure default";

        /// <summary>Reads the grid back into the settings; stays open on a bad value.</summary>
        private void Save()
        {
            grid.EndEdit();

            for (int i = 0; i < settings.Count; i++)
            {
                var setting = settings[i];
                var cells = grid.Rows[i].Cells;
                string displayName = (cells["DisplayName"].Value as string ?? "").Trim();
                string defaultText = (cells["Default"].Value as string ?? "").Trim();

                if (displayName.Length == 0)
                {
                    Problem(i, "DisplayName", $"Enter how {setting.SqlName} is shown.");
                    return;
                }

                if (defaultText.Length > 0 && !IsValid(setting.Parameter, defaultText))
                {
                    Problem(i, "Default", $"\"{defaultText}\" isn't a valid {setting.Parameter.SqlType} for {setting.SqlName}.");
                    return;
                }
            }

            // Check list procedures (and read their choices) before changing anything.
            var lists = new Dictionary<int, ParameterSettings>();

            for (int i = 0; i < settings.Count; i++)
            {
                string listText = (grid.Rows[i].Cells["List"].Value as string ?? "").Trim();

                if (listText.Length == 0)
                    continue;

                var probe = new ParameterSettings(settings[i].Parameter);

                try
                {
                    Cursor = Cursors.WaitCursor;

                    using (var connection = openConnection())
                    {
                        ProcedureList.Apply(probe, listText, connection);
                    }
                }
                catch (Exception ex)
                {
                    Problem(i, "List", $"Choices for {settings[i].SqlName}: {ex.Message}");
                    return;
                }
                finally
                {
                    Cursor = Cursors.Default;
                }

                lists[i] = probe;
            }

            for (int i = 0; i < settings.Count; i++)
            {
                var setting = settings[i];
                var cells = grid.Rows[i].Cells;
                string choice = cells["IfBlank"].Value as string ?? RequiredText;
                string defaultText = (cells["Default"].Value as string ?? "").Trim();

                setting.DisplayName = (cells["DisplayName"].Value as string ?? "").Trim();
                setting.Note = (cells["Note"].Value as string ?? "").Trim();
                setting.DefaultValue = defaultText.Length > 0 ? defaultText : null;
                setting.IfBlank = choice == RequiredText ? BlankValue.Required
                    : choice == NullText ? BlankValue.PassNull
                    : BlankValue.UseProcedureDefault;

                if (lists.TryGetValue(i, out var list))
                {
                    setting.ListProcedure = list.ListProcedure;
                    setting.ListValueColumn = list.ListValueColumn;
                    setting.ListLabelColumn = list.ListLabelColumn;
                    setting.ListRows = list.ListRows;
                }
                else
                {
                    ProcedureList.Clear(setting);
                }
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void Problem(int row, string column, string message)
        {
            grid.CurrentCell = grid.Rows[row].Cells[column];
            MessageBox.Show(this, message, "Flank", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>Text parses for the type; for text types anything goes.</summary>
        internal static bool IsValid(SprocParameter parameter, string text)
        {
            object? value = StoredProcedureRunner.ParseValue(parameter.SqlType, text);

            switch ((parameter.SqlType ?? "").ToLowerInvariant())
            {
                case "tinyint":
                case "smallint":
                case "int":
                case "bigint":
                case "decimal":
                case "numeric":
                case "money":
                case "smallmoney":
                case "float":
                case "real":
                case "bit":
                case "date":
                case "datetime":
                case "datetime2":
                case "smalldatetime":
                    // ParseValue falls back to the text when it can't parse.
                    return !(value is string);

                default:
                    return true;
            }
        }
    }
}
