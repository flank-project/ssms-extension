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
        private const string RequiredText = "Require input";
        private const string NullText = "Pass NULL";

        // For a table-valued parameter.
        private const string RequiredRowText = "Require a row";
        private const string NoRowsText = "Allow no rows (empty table)";

        private readonly IReadOnlyList<ParameterSettings> settings;
        private readonly Func<IDbConnection> openConnection;
        private readonly DataGridView grid;

        // A workbook asks where a table's rows go; a report takes them as a multi-value parameter.
        private readonly bool forWorkbook;

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
            forWorkbook = outputName == "workbook";

            Text = "Flank - " + procedureName;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            ClientSize = new Size(1060, 150 + Math.Min(settings.Count, 10) * 30 + 70);
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
                Size = new Size(1028, ClientSize.Height - 78 - 56),
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
                Name = "Parameter", HeaderText = "Parameter", ReadOnly = true, FillWeight = 18,
                ToolTipText = "The procedure's parameter and its SQL type.",
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "DisplayName", HeaderText = "Shown as", FillWeight = 15,
                ToolTipText = "The label in Excel, or the prompt in the report.",
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Note", HeaderText = "Note", FillWeight = 18,
                ToolTipText = "Help for whoever fills it in (a note on the parameter's label in Excel, or shown when its cell is selected).",
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Initial", HeaderText = "Initial value", FillWeight = 10,
                ToolTipText = "What it's filled in with when the workbook or report opens. Leave empty to start blank.",
            });
            grid.Columns.Add(new DataGridViewComboBoxColumn
            {
                Name = "IfBlank", HeaderText = "If left blank", FillWeight = 27, FlatStyle = FlatStyle.Flat,
                ToolTipText = settings.Any(s => s.Parameter.HasDefault)
                    ? "Require input; omit the parameter from the call, so the procedure's own default applies; or pass NULL."
                    : "Require input, or pass NULL.",
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Guardrails", HeaderText = "Guardrails", ReadOnly = true, FillWeight = 20,
                ToolTipText = "Limits beyond the SQL type, e.g. a dropdown of choices. For a table: where its rows are entered.",
            });
            // Links rather than buttons: an empty one (no guardrail to remove) shows nothing.
            grid.Columns.Add(new DataGridViewLinkColumn
            {
                Name = "EditGuardrails", HeaderText = "", FillWeight = 6, TrackVisitedState = false, LinkBehavior = LinkBehavior.HoverUnderline,
            });
            grid.Columns.Add(new DataGridViewLinkColumn
            {
                Name = "RemoveGuardrails", HeaderText = "", FillWeight = 7, TrackVisitedState = false, LinkBehavior = LinkBehavior.HoverUnderline,
            });

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

                if (parameter.IsTable)
                {
                    // Rows are entered in the workbook or report, not given a starting value here.
                    var row = grid.Rows[i];
                    row.Cells["Parameter"].ToolTipText = "Table columns: " +
                        string.Join(", ", parameter.Columns.Select(c => $"{c.Name} {ProcedureCall.TypeDeclaration(c)}"));
                    row.Cells["Initial"].ReadOnly = true;
                    row.Cells["Initial"].Style.BackColor = SystemColors.Control;
                    row.Cells["Initial"].ToolTipText = "Rows are entered in the " + outputName + ".";

                    choice.Items.Add(RequiredRowText);

                    // A multi-value report parameter needs at least one value.
                    if (forWorkbook)
                        choice.Items.Add(NoRowsText);

                    choice.Value = setting.IfBlank == BlankValue.Required || !forWorkbook ? RequiredRowText : NoRowsText;
                    ShowGuardrails(i);
                    continue;
                }

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

                if (e.ColumnIndex == grid.Columns["EditGuardrails"].Index && settings[e.RowIndex].Parameter.IsTable)
                    EditTableLayout(e.RowIndex);
                else if (e.ColumnIndex == grid.Columns["EditGuardrails"].Index)
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

        // Long defaults are cut short so the column stays narrow.
        private const int DefaultTextLength = 24;

        /// <summary>"Omit (defaults to 5)": in SQL terms, what leaving it out of the EXEC does.</summary>
        private static string UseDefaultText(SprocParameter parameter)
        {
            string text = parameter.DefaultText ?? parameter.DefaultValue ?? "its default";

            if (text.Length > DefaultTextLength)
                text = text.Substring(0, DefaultTextLength - 1).TrimEnd() + "…";

            return $"Omit (defaults to {text})";
        }

        private void ShowGuardrails(int row)
        {
            var setting = settings[row];
            var cells = grid.Rows[row].Cells;

            if (setting.Parameter.IsTable)
            {
                cells["Guardrails"].Value = forWorkbook ? TableLayoutDialog.Summary(setting) : "Multiple values, one per row";
                cells["EditGuardrails"].Value = forWorkbook ? "Edit..." : "";
                cells["RemoveGuardrails"].Value = "";
                return;
            }

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

        private void EditTableLayout(int row)
        {
            if (!forWorkbook)
                return;

            grid.EndEdit();

            var setting = settings[row];
            string typed = (grid.Rows[row].Cells["DisplayName"].Value as string ?? "").Trim();
            if (typed.Length > 0)
                setting.DisplayName = typed;

            using (var dialog = new TableLayoutDialog(setting))
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

                if (initial.Length > 0 && !setting.Parameter.IsTable && !IsValid(setting.Parameter, initial))
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
                setting.DefaultValue = initial.Length > 0 && !setting.Parameter.IsTable ? initial : null;
                setting.IfBlank = choice == RequiredText || choice == RequiredRowText ? BlankValue.Required
                    : choice == NoRowsText ? BlankValue.PassNull
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
