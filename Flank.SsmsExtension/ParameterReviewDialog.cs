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
    /// value it's initialized to, what leaving it blank means, and guardrails
    /// (a dropdown of choices), which open in their own dialog.
    /// </summary>
    internal sealed class ParameterReviewDialog : Form
    {
        private const string RequiredText = "Required";
        private const string NullText = "Pass NULL";

        private readonly IReadOnlyList<ParameterSettings> settings;
        private readonly Func<IDbConnection> openConnection;
        private readonly DataGridView grid;

        /// <param name="outputName">What's being created, e.g. "workbook".</param>
        /// <param name="openConnection">Opens a connection to the database, to check list procedures.</param>
        public ParameterReviewDialog(
            string procedureName,
            IReadOnlyList<ParameterSettings> settings,
            string outputName,
            Func<IDbConnection> openConnection)
        {
            this.settings = settings;
            this.openConnection = openConnection;

            Text = "Flank - " + procedureName;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            ClientSize = new Size(980, 150 + Math.Min(settings.Count, 10) * 30 + 70);
            MinimumSize = new Size(760, 320);

            Controls.Add(new Label
            {
                Text = "Parameters",
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 16, FontStyle.Bold),
                Location = new Point(14, 12),
                AutoSize = true,
            });

            Controls.Add(new Label
            {
                Text = $"How people using the {outputName} see and fill in each parameter of {procedureName}.",
                ForeColor = SystemColors.GrayText,
                Location = new Point(16, 48),
                AutoSize = true,
            });

            grid = new DataGridView
            {
                Location = new Point(16, 78),
                Size = new Size(948, ClientSize.Height - 78 - 56),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                EditMode = DataGridViewEditMode.EditOnEnter,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            };

            grid.RowTemplate.Height = 28;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold);

            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Parameter", HeaderText = "Parameter", ReadOnly = true, FillWeight = 20,
                ToolTipText = "The procedure's parameter and its SQL type.",
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "DisplayName", HeaderText = "Shown as", FillWeight = 16,
                ToolTipText = "The label in Excel, or the prompt in the report.",
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Note", HeaderText = "Note", FillWeight = 22,
                ToolTipText = "Help for whoever fills it in (shown next to the Excel cell).",
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Initial", HeaderText = "Initial value", FillWeight = 12,
                ToolTipText = "What it's filled in with when the workbook or report opens. Leave empty to start blank.",
            });
            grid.Columns.Add(new DataGridViewComboBoxColumn
            {
                Name = "IfBlank", HeaderText = "If left blank", FillWeight = 18, FlatStyle = FlatStyle.Flat,
                ToolTipText = "Required; use the procedure's own default (the parameter is left out of the call); or pass NULL.",
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Guardrails", HeaderText = "Guardrails", ReadOnly = true, FillWeight = 22,
                ToolTipText = "Limits beyond the SQL type, e.g. a dropdown of choices from a procedure.",
            });
            grid.Columns.Add(new DataGridViewButtonColumn { Name = "EditGuardrails", HeaderText = "", FillWeight = 7, FlatStyle = FlatStyle.System });
            grid.Columns.Add(new DataGridViewButtonColumn { Name = "RemoveGuardrails", HeaderText = "", FillWeight = 7, FlatStyle = FlatStyle.System });

            grid.Columns["Parameter"].DefaultCellStyle.ForeColor = SystemColors.GrayText;
            grid.Columns["Guardrails"].DefaultCellStyle.ForeColor = SystemColors.GrayText;

            for (int i = 0; i < settings.Count; i++)
            {
                var setting = settings[i];
                var parameter = setting.Parameter;

                grid.Rows.Add(
                    $"{parameter.Name} {ProcedureCall.TypeDeclaration(parameter)}",
                    setting.DisplayName,
                    setting.Note,
                    setting.DefaultValue ?? "",
                    null);

                var choice = (DataGridViewComboBoxCell)grid.Rows[i].Cells["IfBlank"];
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

                ShowGuardrails(i);
            }

            grid.CellContentClick += (s, e) =>
            {
                if (e.RowIndex < 0)
                    return;

                if (e.ColumnIndex == grid.Columns["EditGuardrails"].Index)
                    EditGuardrails(e.RowIndex);
                else if (e.ColumnIndex == grid.Columns["RemoveGuardrails"].Index && settings[e.RowIndex].HasList)
                {
                    ProcedureList.Clear(settings[e.RowIndex]);
                    ShowGuardrails(e.RowIndex);
                }
            };

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

        private void ShowGuardrails(int row)
        {
            var setting = settings[row];
            var cells = grid.Rows[row].Cells;

            cells["Guardrails"].Value = setting.HasList ? GuardrailsDialog.Summary(setting) : "None";
            cells["EditGuardrails"].Value = setting.HasList ? "Edit..." : "Add...";
            cells["RemoveGuardrails"].Value = setting.HasList ? "Remove" : "";
        }

        private void EditGuardrails(int row)
        {
            grid.EndEdit();

            // The dialog's heading uses the name as currently typed.
            var setting = settings[row];
            string typed = (grid.Rows[row].Cells["DisplayName"].Value as string ?? "").Trim();
            if (typed.Length > 0)
                setting.DisplayName = typed;

            using (var dialog = new GuardrailsDialog(setting, openConnection))
            {
                dialog.ShowDialog(this);
            }

            ShowGuardrails(row);
        }

        /// <summary>Reads the grid back into the settings; stays open on a bad value.</summary>
        private void Save()
        {
            grid.EndEdit();

            for (int i = 0; i < settings.Count; i++)
            {
                var setting = settings[i];
                var cells = grid.Rows[i].Cells;
                string displayName = (cells["DisplayName"].Value as string ?? "").Trim();
                string initial = (cells["Initial"].Value as string ?? "").Trim();

                if (displayName.Length == 0)
                {
                    Problem(i, "DisplayName", $"Enter how {setting.SqlName} is shown.");
                    return;
                }

                if (initial.Length > 0 && !IsValid(setting.Parameter, initial))
                {
                    Problem(i, "Initial", $"\"{initial}\" isn't a valid {setting.Parameter.SqlType} for {setting.SqlName}.");
                    return;
                }
            }

            for (int i = 0; i < settings.Count; i++)
            {
                var setting = settings[i];
                var cells = grid.Rows[i].Cells;
                string choice = cells["IfBlank"].Value as string ?? RequiredText;
                string initial = (cells["Initial"].Value as string ?? "").Trim();

                setting.DisplayName = (cells["DisplayName"].Value as string ?? "").Trim();
                setting.Note = (cells["Note"].Value as string ?? "").Trim();
                setting.DefaultValue = initial.Length > 0 ? initial : null;
                setting.IfBlank = choice == RequiredText ? BlankValue.Required
                    : choice == NullText ? BlankValue.PassNull
                    : BlankValue.UseProcedureDefault;
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
